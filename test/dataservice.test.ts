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

// Run in a zone that is not UTC, as CI runners usually are, so a timestamp converted to UTC
// fails. A half-hour offset also catches an off-by-whole-hours conversion.
process.env.TZ = 'Asia/Kolkata';

const FIXTURES = path.join(__dirname, 'fixtures');
const fixture = (name: string) => path.join(FIXTURES, name);

// Timestamps are sent as ISO 8601 local wall-clock time with no offset, whatever the machine's
// locale, so expectations are literal strings.

function writeTemp(name: string, content: string): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'worklogs-'));
  const file = path.join(dir, name);
  fs.writeFileSync(file, content);
  return file;
}

async function parseErrors(fn: () => Promise<unknown>): Promise<string[]> {
  try {
    await fn();
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
    test(`maps every configured column (${file})`, async () => {
      const result = await service().getMappingData(fixture(file), false, API);

      assert.deepEqual([...result.keys()], [API.toLowerCase()]);
      assert.deepEqual(result.get(API.toLowerCase()), [
        {
          userId: 'alice@example.com',
          timeStamp: '2026-10-01T09:30:00',
          length: 5400,
          billableLength: 3600,
          activityTypeId: 'Development',
          comment: 'Fixed bug',
          workItemId: 101,
        },
        {
          userId: 'bob@example.com',
          timeStamp: '2026-10-02T14:00:00',
          length: 900,
          billableLength: 0,
          activityTypeId: 'Meeting',
          comment: 'Standup',
        },
      ]);
    });
  }

  test('timestamps do not depend on the default locale', async (t) => {
    // Node's default locale comes from the OS, so simulate a German one: any call that
    // omits the locale argument now formats as de-DE.
    const original = Date.prototype.toLocaleString;
    t.mock.method(Date.prototype, 'toLocaleString', function (this: Date, locales?: string, options?: Intl.DateTimeFormatOptions) {
      return original.call(this, locales ?? 'de-DE', options);
    });
    const [log] = (await service().getMappingData(fixture('positional.csv'), false, API)).get(API.toLowerCase())!;
    assert.equal(log.timeStamp, '2026-10-01T09:30:00');
  });

  test('reads only the first sheet and skips blank rows', async () => {
    const logs = (await service().getMappingData(fixture('positional.xlsx'), false, API)).get(API.toLowerCase())!;
    assert.equal(logs.length, 2);
    assert.ok(logs.every(w => w.userId !== 'should-not-be-read'));
  });

  test('accepts a comma as the decimal separator and rounds seconds up', async () => {
    const file = writeTemp('w.csv', 'a@example.com,2026-10-01 09:30,"1,5",,c,1,"0,0001"\n');
    const [log] = (await service().getMappingData(file, false, API)).get(API.toLowerCase())!;
    assert.equal(log.length, 5400);
    assert.equal(log.billableLength, 1);
  });

  test('a row with only a comment needs no work item', async () => {
    const file = writeTemp('w.csv', ',,1,,just a comment,,\n');
    const [log] = (await service().getMappingData(file, false, API)).get(API.toLowerCase())!;
    assert.deepEqual(log, { length: 3600, comment: 'just a comment' });
  });

  test('an empty file yields no organizations', async () => {
    const file = writeTemp('w.csv', '');
    assert.equal((await service().getMappingData(file, false, API)).size, 0);
  });

  for (const bad of ['0', '-3', '1.5', 'abc']) {
    test(`rejects work item "${bad}" naming the cell`, async () => {
      const file = writeTemp('w.csv', `a,,1,,c,1\na,,1,,c,${bad}\n`);
      const errors = await parseErrors(() => service().getMappingData(file, false, API));
      assert.deepEqual(errors, [
        `"WorkItem" is not a valid value in a cell F2. The value may be not specified or must be an integer and greater than 0.`,
      ]);
    });
  }

  test('rejects an unparseable timestamp', async () => {
    const file = writeTemp('w.csv', 'a,not a date,1,,c,1\n');
    const errors = await parseErrors(() => service().getMappingData(file, false, API));
    assert.deepEqual(errors, ['"TimeStamp" is not a valid datetime value in a cell B1.']);
  });

  for (const bad of ['0', '-1', 'abc']) {
    test(`rejects duration "${bad}"`, async () => {
      const file = writeTemp('w.csv', `a,,${bad},,c,1\n`);
      const errors = await parseErrors(() => service().getMappingData(file, false, API));
      assert.deepEqual(errors, [
        `"Duration" is not a valid value in a cell C1. The value may be not specified or must be greater than 0.`,
      ]);
    });
  }

  test('rejects a negative billable duration', async () => {
    const file = writeTemp('w.csv', 'a,,1,,c,1,-1\n');
    const errors = await parseErrors(() => service().getMappingData(file, false, API));
    assert.deepEqual(errors, [
      `"BillableDuration" is not a valid value in a cell G1. The value may be not specified or must be greater or equal than 0.`,
    ]);
  });

  test('requires a comment or a work item', async () => {
    const file = writeTemp('w.csv', 'a,,1,,,,\n');
    const errors = await parseErrors(() => service().getMappingData(file, false, API));
    assert.deepEqual(errors, [
      `"Comment" (cell E1) and "WorkItem" (cell F1) are not specified. At least one of these values must be specified.`,
    ]);
  });

  test('collects errors from every row before throwing', async () => {
    const file = writeTemp('w.csv', 'a,,0,,c,1\na,,1,,c,x\n');
    const errors = await parseErrors(() => service().getMappingData(file, false, API));
    assert.equal(errors.length, 2);
    assert.match(errors[0], /"Duration".*C1/);
    assert.match(errors[1], /"WorkItem".*F2/);
  });

  test('without an API url, rows group under the default organization, lowercased', async () => {
    const result = await service().getMappingData(fixture('positional.csv'), false, undefined, 'MyOrg');
    assert.deepEqual([...result.keys()], ['myorg']);
    assert.equal(result.get('myorg')!.length, 2);
  });

  test('without an API url or a default organization, every row is an error', async () => {
    for (const org of [undefined, '   ']) {
      const errors = await parseErrors(() => service().getMappingData(fixture('positional.csv'), false, undefined, org));
      assert.equal(errors.length, 2);
      assert.ok(errors.every(e => e.includes('"Organization"')));
    }
  });
});

describe('getMappingData, column names from the header row (-m)', () => {
  for (const file of ['headers.xlsx', 'headers.csv']) {
    test(`groups by the Organization column and combines Date with Start (${file})`, async () => {
      const result = await service().getMappingData(fixture(file), true, undefined, 'Fallback');

      assert.deepEqual([...result.keys()], ['contosoa', 'fabrikam']);
      assert.deepEqual(result.get('contosoa'), [
        {
          userId: 'alice@example.com',
          workItemId: 101,
          timeStamp: '2026-10-01T09:30:00',
          length: 7200,
          billableLength: 5400,
          activityTypeId: 'Development',
          comment: 'Review',
        },
        {
          userId: 'bob@example.com',
          workItemId: 102,
          timeStamp: '2026-10-02T00:00:00',
          length: 1800,
          comment: 'Planning',
        },
      ]);
      assert.deepEqual(result.get('fabrikam'), [
        {
          userId: 'carol@example.com',
          timeStamp: '2026-10-03T13:15:00',
          length: 3600,
          comment: 'Support',
        },
      ]);
    });
  }

  test('matches headers in a CSV saved with a UTF-8 byte-order mark (Excel "CSV UTF-8")', async () => {
    const file = writeTemp('w.csv', '﻿UserName,WorkItem,Duration,Comment\na,1,1,c\n');
    const [log] = (await service().getMappingData(file, true, 'https://x')).get('https://x')!;
    assert.deepEqual(log, { userId: 'a', workItemId: 1, length: 3600, comment: 'c' });
  });

  test('skips rows whose fields are all empty', async () => {
    const file = writeTemp('w.csv', 'WorkItem,Duration,Comment\n1,1,c\n,,\n2,1,d\n');
    assert.equal((await service().getMappingData(file, true, 'https://x')).get('https://x')!.length, 2);
  });

  test('an Organization cell wins over the API url', async () => {
    const result = await service().getMappingData(fixture('headers.csv'), true, 'https://ignored.example.com');
    assert.deepEqual([...result.keys()], ['contosoa', 'fabrikam']);
  });

  test('a row without an Organization falls back to the default organization', async () => {
    const file = writeTemp('w.csv', 'UserName,WorkItem,Duration,Comment,Organization\na,1,1,c,\n');
    const result = await service().getMappingData(file, true, undefined, 'Fallback');
    assert.deepEqual([...result.keys()], ['fallback']);
  });

  test('error cells are numbered from the data row under the header', async () => {
    const file = writeTemp('w.csv', 'WorkItem,Duration,Comment\n1,1,c\nx,1,c\n');
    const errors = await parseErrors(() => service().getMappingData(file, true, 'https://x'));
    assert.deepEqual(errors, [
      `"WorkItem" is not a valid value in a cell WorkItem3. The value may be not specified or must be an integer and greater than 0.`,
    ]);
  });

  test('rejects a Date and Start that do not form a datetime', async () => {
    const file = writeTemp('w.csv', 'Date,Start,Duration,Comment\nnope,9:00 AM,1,c\n');
    const errors = await parseErrors(() => service().getMappingData(file, true, 'https://x'));
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
