import assert from 'assert';
import { networkAccess } from './networkAccess.js';

const env = Object.assign({}, process.env);

describe('utils/networkAccess', () => {
  afterEach(() => {
    process.env = Object.assign({}, env);
  });

  it('restricts external access by default', () => {
    delete process.env.CLIMICROSOFT365_ALLOW_EXTERNAL;
    delete process.env.CLIMICROSOFT365_RESTRICT_EXTERNAL;

    assert.strictEqual(networkAccess.isRestricted(), true);
  });

  it('allows overriding the restriction explicitly', () => {
    process.env.CLIMICROSOFT365_ALLOW_EXTERNAL = '1';

    assert.strictEqual(networkAccess.isRestricted(), false);
  });

  it('allows Microsoft 365 hosts in restricted mode', () => {
    process.env.CLIMICROSOFT365_RESTRICT_EXTERNAL = '1';

    assert.doesNotThrow(() => networkAccess.assertAllowed('https://graph.microsoft.com/v1.0/me'));
    assert.doesNotThrow(() => networkAccess.assertAllowed('https://contoso.sharepoint.com/sites/demo'));
  });

  it('blocks non Microsoft 365 hosts in restricted mode', () => {
    process.env.CLIMICROSOFT365_RESTRICT_EXTERNAL = '1';

    assert.throws(
      () => networkAccess.assertAllowed('https://haveibeenpwned.com/api/v3/breachedaccount/example@contoso.com'),
      /blocked by default/
    );
  });

  it('allows explicitly allowlisted hosts in restricted mode', () => {
    process.env.CLIMICROSOFT365_RESTRICT_EXTERNAL = '1';
    process.env.CLIMICROSOFT365_ALLOWED_HOSTS = 'haveibeenpwned.com';

    assert.doesNotThrow(() => networkAccess.assertAllowed('https://haveibeenpwned.com/api/v3/breachedaccount/example@contoso.com'));
  });
});
