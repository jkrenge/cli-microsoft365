import assert from 'assert';
import fs from 'fs';
import os from 'os';
import path from 'path';
import sinon from 'sinon';
import { mailSenderWhitelist } from './mailSenderWhitelist.js';

describe('mailSenderWhitelist', () => {
  let tempHome: string;

  beforeEach(() => {
    tempHome = fs.mkdtempSync(path.join(os.tmpdir(), 'cli-m365-mail-whitelist-'));
    sinon.stub(os, 'homedir').returns(tempHome);
  });

  afterEach(() => {
    sinon.restore();
    fs.rmSync(tempHome, { force: true, recursive: true });
  });

  it('returns an empty store when the whitelist file does not exist', () => {
    assert.deepStrictEqual(mailSenderWhitelist.load(), {
      domains: [],
      senders: []
    });
  });

  it('stores normalized sender and domain entries in the home directory', () => {
    assert.strictEqual(mailSenderWhitelist.addSender(' Allowed@Example.com '), true);
    assert.strictEqual(mailSenderWhitelist.addDomain(' Contoso.com '), true);
    assert.strictEqual(mailSenderWhitelist.addSender('allowed@example.com'), false);
    assert.strictEqual(mailSenderWhitelist.addDomain('contoso.com'), false);

    assert.deepStrictEqual(JSON.parse(fs.readFileSync(mailSenderWhitelist.getFilePath(), 'utf8')), {
      domains: ['contoso.com'],
      senders: ['allowed@example.com']
    });
  });

  it('filters messages using sender and domain matches', () => {
    fs.writeFileSync(mailSenderWhitelist.getFilePath(), JSON.stringify({
      domains: ['contoso.com'],
      senders: ['allowed@example.com']
    }), 'utf8');

    const filterResult = mailSenderWhitelist.filterMessages([
      {
        from: {
          emailAddress: {
            address: 'allowed@example.com'
          }
        },
        subject: 'allowed sender'
      },
      {
        from: {
          emailAddress: {
            address: 'worker@contoso.com'
          }
        },
        subject: 'allowed domain'
      },
      {
        from: {
          emailAddress: {
            address: 'unknown@fabrikam.com'
          }
        },
        subject: 'blocked'
      }
    ]);

    assert.strictEqual(filterResult.allowed.length, 2);
    assert.strictEqual(filterResult.blocked.length, 1);
    assert.strictEqual(filterResult.filteredCount, 1);
  });

  it('removes nextLink when a collection contains only blocked senders', () => {
    const response = {
      '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/messages?$skiptoken=abc',
      value: [
        {
          from: {
            emailAddress: {
              address: 'unknown@fabrikam.com'
            }
          },
          subject: 'blocked'
        }
      ]
    };

    const filterResult = mailSenderWhitelist.filterResponse(response);

    assert.deepStrictEqual(filterResult.filteredResponse, {
      value: []
    });
    assert.strictEqual(filterResult.filteredCount, 1);
  });
});
