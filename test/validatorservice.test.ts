import 'reflect-metadata';
import { test, describe, beforeEach, afterEach } from 'node:test';
import * as assert from 'node:assert/strict';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { ValidatorService } from '../src/services/validatorservice';

let dir: string;
let printed: string[];
const originalError = console.error;

beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'validator-'));
  printed = [];
  console.error = (msg: string) => { printed.push(String(msg)); };
});
afterEach(() => { console.error = originalError; });

function file(name: string): string {
  const p = path.join(dir, name);
  fs.writeFileSync(p, '');
  return p;
}

// A token-authorized run that passes every check; each test overrides one field.
const valid = () => ({ file: file('worklogs.xlsx'), authorization: 'token', token: 't' });
const validate = (program: object) => new ValidatorService().validate(program);

describe('validate', () => {
  test('accepts a complete token run', () => {
    assert.equal(validate(valid()), false);
    assert.deepEqual(printed, []);
  });

  for (const name of ['w.xlsx', 'w.csv', 'W.XLSX', 'W.Csv']) {
    test(`accepts ${name}`, () => {
      assert.equal(validate({ ...valid(), file: file(name) }), false);
    });
  }

  for (const name of ['w.xls', 'w.txt', 'w.json', 'xlsx']) {
    test(`rejects ${name}`, () => {
      assert.equal(validate({ ...valid(), file: file(name) }), true);
      assert.match(printed.join('\n'), /is not in supported format/);
    });
  }

  test('rejects a missing --file', () => {
    assert.equal(validate({ ...valid(), file: undefined }), true);
    assert.match(printed.join('\n'), /File is not specified/);
  });

  test('rejects a file that does not exist', () => {
    assert.equal(validate({ ...valid(), file: path.join(dir, 'absent.xlsx') }), true);
    assert.match(printed.join('\n'), /does not exist/);
  });

  test('rejects an unknown authorization type', () => {
    assert.equal(validate({ ...valid(), authorization: 'basic' }), true);
    assert.match(printed.join('\n'), /Unknown API authorization type/);
  });

  test('accepts ntlm with api, organization, user and password', () => {
    const program = { file: file('w.csv'), authorization: 'NTLM', api: 'https://x', organization: 'o', user: 'u', password: 'p' };
    assert.equal(validate(program), false);
  });

  test('ntlm requires api and organization', () => {
    const program = { file: file('w.csv'), authorization: 'ntlm', user: 'u', password: 'p' };
    assert.equal(validate(program), true);
    assert.match(printed.join('\n'), /Api url and organization are not specified/);
  });

  test('ntlm requires user and password', () => {
    const program = { file: file('w.csv'), authorization: 'ntlm', api: 'https://x', organization: 'o' };
    assert.equal(validate(program), true);
    assert.match(printed.join('\n'), /User and password are not specified/);
  });
});
