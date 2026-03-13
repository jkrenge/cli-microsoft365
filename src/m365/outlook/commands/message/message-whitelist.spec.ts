import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { mailSenderWhitelist } from '../../../../utils/mailSenderWhitelist.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './message-whitelist.js';

describe(commands.MESSAGE_WHITELIST, () => {
  const emailResponse = {
    value: [
      {
        id: '1',
        receivedDateTime: '2026-03-13T19:22:47Z',
        subject: 'blocked',
        from: {
          emailAddress: {
            address: 'unknown@fabrikam.com'
          }
        }
      }
    ]
  };

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.accessTokens[auth.defaultResource] = {
      accessToken: 'abc',
      expiresOn: 'abc'
    };
    commandInfo = cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: any) => {
        log.push(msg);
      },
      logRaw: async (msg: any) => {
        log.push(msg);
      },
      logToStderr: async (msg: any) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
  });

  afterEach(() => {
    sinonUtil.restore([
      accessToken.isAppOnlyAccessToken,
      cli.promptForInput,
      mailSenderWhitelist.addDomain,
      mailSenderWhitelist.filterMessages,
      mailSenderWhitelist.getFilePath,
      mailSenderWhitelist.getSenderAddress,
      mailSenderWhitelist.getSenderDomain,
      mailSenderWhitelist.shouldFilterInInteractiveMode,
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.accessTokens = {};
  });

  it('fails validation when interactive mode is disabled', async () => {
    sinon.stub(mailSenderWhitelist, 'shouldFilterInInteractiveMode').returns(false);

    const actual = await command.validate({
      options: {}
    }, commandInfo);

    assert.strictEqual(actual, 'This command requires interactive mode. Enable prompts and run it from a terminal session.');
  });

  it('whitelists a sender domain interactively', async () => {
    sinon.stub(mailSenderWhitelist, 'shouldFilterInInteractiveMode').returns(true);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/messages?$top=25&$orderby=receivedDateTime desc') {
        return emailResponse;
      }

      throw 'Invalid request';
    });
    sinon.stub(mailSenderWhitelist, 'filterMessages')
      .onFirstCall()
      .returns({
        allowed: [],
        blocked: emailResponse.value,
        filteredCount: 1
      })
      .onSecondCall()
      .returns({
        allowed: emailResponse.value,
        blocked: [],
        filteredCount: 0
      });
    sinon.stub(mailSenderWhitelist, 'getSenderAddress').returns('unknown@fabrikam.com');
    sinon.stub(mailSenderWhitelist, 'getSenderDomain').returns('fabrikam.com');
    sinon.stub(mailSenderWhitelist, 'addDomain').returns(true);
    sinon.stub(mailSenderWhitelist, 'getFilePath').returns('/Users/julian/.cli-m365-mail-sender-whitelist.json');
    sinon.stub(cli, 'promptForInput').resolves('a');

    await command.action(logger, { options: {} } as any);

    assert(loggerLogToStderrSpy.calledWith('Sender: unknown@fabrikam.com'));
    assert(loggerLogToStderrSpy.calledWith('Subject: blocked'));
    assert(loggerLogToStderrSpy.calledWith('Whitelisted domain fabrikam.com'));
    assert(loggerLogSpy.calledWith({
      domainsAdded: 1,
      filePath: '/Users/julian/.cli-m365-mail-sender-whitelist.json',
      remaining: 0,
      sendersAdded: 0,
      skipped: 0
    }));
  });
});
