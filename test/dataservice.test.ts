import 'reflect-metadata';
import { test, describe } from 'node:test';
import * as assert from 'node:assert/strict';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { DataService } from '../src/services/dataservice';
import { APIService } from '../src/services/apiservice';
import { AuthType } from '../src/models/authtype';
import { Worklog } from '../src/models/worklog';

const FIXTURES = path.join(__dirname, 'fixtures');
const fixture = (name: string) => path.join(FIXTURES, name);

// The service renders timestamps with toLocaleString(), so expectations are built the same
// way from a local wall-clock time: the assertions hold in any locale and time zone.
const local = (y: number, m: number, d: number, hh = 0, mm = 0) =>
  new Date(y, m - 1, d, hh, mm).toLocaleString();

function writeTemp(name: string, content: string): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'worklogs-'));
  const file = path.join(dir, name);
  fs.writeFileSync(file, content);
  return file;
}

function parseErrors(fn: () => unknown): string[] {
  try {
    fn();
  } catch (e) {
    assert.ok(Array.isArray(e), 'errors are thrown as an array of messages');
    return e as string[];
  }
  assert.fail('expected getMappingData to throw');
}

const service = (api: Partial<APIService> = {}) => new DataService(api as APIService);

describe('getMappingData, positional columns from config.json', () => {
  const API = 'https://Contoso.timehub.7pace.com';

  for (const file of ['positional.xlsx', 'positional.csv']) {
    test(`maps every configured column (${file})`, () => {
      const result = service().getMappingData(fixture(file), false, API);

      assert.deepEqual([...result.keys()], [API.toLowerCase()]);
      assert.deepEqual(result.get(API.toLowerCase()), [
        {
          userId: 'alice@example.com',
          timeStamp: local(2026, 10, 1, 9, 30),
          length: 5400,
          billableLength: 3600,
          activityTypeId: 'Development',
          comment: 'Fixed bug',
          workItemId: 101,
        },
        {
          userId: 'bob@example.com',
          timeStamp: local(2026, 10, 2, 14, 0),
          length: 900,
          billableLength: 0,
          activityTypeId: 'Meeting',
          comment: 'Standup',
        },
      ]);
    });
  }

  test('reads only the first sheet and skips blank rows', () => {
    const logs = service().getMappingData(fixture('positional.xlsx'), false, API).get(API.toLowerCase())!;
    assert.equal(logs.length, 2);
    assert.ok(logs.every(w => w.userId !== 'should-not-be-read'));
  });

  test('accepts a comma as the decimal separator and rounds seconds up', () => {
    const file = writeTemp('w.csv', 'a@example.com,2026-10-01 09:30,"1,5",,c,1,"0,0001"\n');
    const [log] = service().getMappingData(file, false, API).get(API.toLowerCase())!;
    assert.equal(log.length, 5400);
    assert.equal(log.billableLength, 1);
  });

  test('a row with only a comment needs no work item', () => {
    const file = writeTemp('w.csv', ',,1,,just a comment,,\n');
    const [log] = service().getMappingData(file, false, API).get(API.toLowerCase())!;
    assert.deepEqual(log, { length: 3600, comment: 'just a comment' });
  });

  test('an empty file yields no organizations', () => {
    const file = writeTemp('w.csv', '');
    assert.equal(service().getMappingData(file, false, API).size, 0);
  });

  for (const bad of ['0', '-3', '1.5', 'abc']) {
    test(`rejects work item "${bad}" naming the cell`, () => {
      const file = writeTemp('w.csv', `a,,1,,c,1\na,,1,,c,${bad}\n`);
      const errors = parseErrors(() => service().getMappingData(file, false, API));
      assert.deepEqual(errors, [
        `"WorkItem" is not a valid value in a cell F2. The value may be not specified or must be an integer and greater than 0.`,
      ]);
    });
  }

  test('rejects an unparseable timestamp', () => {
    const file = writeTemp('w.csv', 'a,not a date,1,,c,1\n');
    const errors = parseErrors(() => service().getMappingData(file, false, API));
    assert.deepEqual(errors, ['"TimeStamp" is not a valid datetime value in a cell B1.']);
  });

  for (const bad of ['0', '-1', 'abc']) {
    test(`rejects duration "${bad}"`, () => {
      const file = writeTemp('w.csv', `a,,${bad},,c,1\n`);
      const errors = parseErrors(() => service().getMappingData(file, false, API));
      assert.deepEqual(errors, [
        `"Duration" is not a valid value in a cell C1. The value may be not specified or must be greater than 0.`,
      ]);
    });
  }

  test('rejects a negative billable duration', () => {
    const file = writeTemp('w.csv', 'a,,1,,c,1,-1\n');
    const errors = parseErrors(() => service().getMappingData(file, false, API));
    assert.deepEqual(errors, [
      `"BillableDuration" is not a valid value in a cell G1. The value may be not specified or must be greater or equal than 0.`,
    ]);
  });

  test('requires a comment or a work item', () => {
    const file = writeTemp('w.csv', 'a,,1,,,,\n');
    const errors = parseErrors(() => service().getMappingData(file, false, API));
    assert.deepEqual(errors, [
      `"Comment" (cell E1) and "WorkItem" (cell F1) are not specified. At least one of these values must be specified.`,
    ]);
  });

  test('collects errors from every row before throwing', () => {
    const file = writeTemp('w.csv', 'a,,0,,c,1\na,,1,,c,x\n');
    const errors = parseErrors(() => service().getMappingData(file, false, API));
    assert.equal(errors.length, 2);
    assert.match(errors[0], /"Duration".*C1/);
    assert.match(errors[1], /"WorkItem".*F2/);
  });

  test('without an API url, rows group under the default organization, lowercased', () => {
    const result = service().getMappingData(fixture('positional.csv'), false, undefined, 'MyOrg');
    assert.deepEqual([...result.keys()], ['myorg']);
    assert.equal(result.get('myorg')!.length, 2);
  });

  test('without an API url or a default organization, every row is an error', () => {
    for (const org of [undefined, '   ']) {
      const errors = parseErrors(() => service().getMappingData(fixture('positional.csv'), false, undefined, org));
      assert.equal(errors.length, 2);
      assert.ok(errors.every(e => e.includes('"Organization"')));
    }
  });
});

describe('getMappingData, column names from the header row (-m)', () => {
  for (const file of ['headers.xlsx', 'headers.csv']) {
    test(`groups by the Organization column and combines Date with Start (${file})`, () => {
      const result = service().getMappingData(fixture(file), true, undefined, 'Fallback');

      assert.deepEqual([...result.keys()], ['contosoa', 'fabrikam']);
      assert.deepEqual(result.get('contosoa'), [
        {
          userId: 'alice@example.com',
          workItemId: 101,
          timeStamp: local(2026, 10, 1, 9, 30),
          length: 7200,
          billableLength: 5400,
          activityTypeId: 'Development',
          comment: 'Review',
        },
        {
          userId: 'bob@example.com',
          workItemId: 102,
          timeStamp: local(2026, 10, 2),
          length: 1800,
          comment: 'Planning',
        },
      ]);
      assert.deepEqual(result.get('fabrikam'), [
        {
          userId: 'carol@example.com',
          timeStamp: local(2026, 10, 3, 13, 15),
          length: 3600,
          comment: 'Support',
        },
      ]);
    });
  }

  test('an Organization cell wins over the API url', () => {
    const result = service().getMappingData(fixture('headers.csv'), true, 'https://ignored.example.com');
    assert.deepEqual([...result.keys()], ['contosoa', 'fabrikam']);
  });

  test('a row without an Organization falls back to the default organization', () => {
    const file = writeTemp('w.csv', 'UserName,WorkItem,Duration,Comment,Organization\na,1,1,c,\n');
    const result = service().getMappingData(file, true, undefined, 'Fallback');
    assert.deepEqual([...result.keys()], ['fallback']);
  });

  test('error cells are numbered from the data row under the header', () => {
    const file = writeTemp('w.csv', 'WorkItem,Duration,Comment\n1,1,c\nx,1,c\n');
    const errors = parseErrors(() => service().getMappingData(file, true, 'https://x'));
    assert.deepEqual(errors, [
      `"WorkItem" is not a valid value in a cell WorkItem3. The value may be not specified or must be an integer and greater than 0.`,
    ]);
  });

  test('rejects a Date and Start that do not form a datetime', () => {
    const file = writeTemp('w.csv', 'Date,Start,Duration,Comment\nnope,9:00 AM,1,c\n');
    const errors = parseErrors(() => service().getMappingData(file, true, 'https://x'));
    assert.deepEqual(errors, [
      `"Date" and "Start" do not specify a valid starting datetime value "nope 9:00 AM" in a cell(s) Date2.`,
    ]);
  });
});

describe('getUserIdActivityTypeId', () => {
  const args = ['https://x', AuthType.token, 'tok', 'org', 'u', 'p'] as const;

  function fakeApi(calls: string[]) {
    return {
      getUsers: async () => { calls.push('users'); return new Map([['alice@example.com', 'u-1']]); },
      getActivityTypes: async () => { calls.push('types'); return new Map([['Development', 'a-1']]); },
    };
  }

  test('replaces user names and activity type names with their ids', async () => {
    const calls: string[] = [];
    const logs: Worklog[] = [{ userId: 'alice@example.com', activityTypeId: 'Development', comment: 'c' }];
    const out = await service(fakeApi(calls)).getUserIdActivityTypeId(logs, ...args);
    assert.deepEqual(out, [{ userId: 'u-1', activityTypeId: 'a-1', comment: 'c' }]);
    assert.deepEqual(calls, ['users', 'types']);
  });

  test('skips the lookups no worklog needs', async () => {
    const calls: string[] = [];
    await service(fakeApi(calls)).getUserIdActivityTypeId([{ comment: 'c' }], ...args);
    assert.deepEqual(calls, []);
  });

  test('reports each unknown name once', async () => {
    const logs: Worklog[] = [
      { userId: 'ghost', activityTypeId: 'Nope' },
      { userId: 'ghost', activityTypeId: 'Nope' },
    ];
    await assert.rejects(service(fakeApi([])).getUserIdActivityTypeId(logs, ...args), (e: unknown) => {
      assert.deepEqual(e, [
        'User "ghost" does not exist in your organization',
        'Activity type "Nope" does not exist in your organization',
      ]);
      return true;
    });
  });
});

describe('sendWorklogs', () => {
  test('posts a validation batch, then the real batch', async () => {
    const calls: boolean[] = [];
    const api = { postWorklogsBatch: async (_: Worklog[], validateOnly: boolean) => { calls.push(validateOnly); } };
    await service(api).sendWorklogs([{ comment: 'c' }], 'https://x', AuthType.token, 't', 'o', 'u', 'p');
    assert.deepEqual(calls, [true, false]);
  });
});
