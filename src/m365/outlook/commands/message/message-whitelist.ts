import { Message } from '@microsoft/microsoft-graph-types';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { mailSenderWhitelist } from '../../../../utils/mailSenderWhitelist.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { Outlook } from '../../Outlook.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  folderId?: string;
  folderName?: string;
  startTime?: string;
  endTime?: string;
  top?: number;
  userId?: string;
  userName?: string;
}

type ReviewAction = 'a' | 's' | 'd';

class OutlookMessageWhitelistCommand extends GraphCommand {
  public get name(): string {
    return commands.MESSAGE_WHITELIST;
  }

  public get description(): string {
    return 'Interactively review and whitelist mail senders';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        folderId: typeof args.options.folderId !== 'undefined',
        folderName: typeof args.options.folderName !== 'undefined',
        startTime: typeof args.options.startTime !== 'undefined',
        endTime: typeof args.options.endTime !== 'undefined',
        top: args.options.top || 25,
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--folderName [folderName]',
        autocomplete: Outlook.wellKnownFolderNames
      },
      {
        option: '--folderId [folderId]'
      },
      {
        option: '--startTime [startTime]'
      },
      {
        option: '--endTime [endTime]'
      },
      {
        option: '--top [top]'
      },
      {
        option: '--userId [userId]'
      },
      {
        option: '--userName [userName]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!mailSenderWhitelist.shouldFilterInInteractiveMode()) {
          return 'This command requires interactive mode. Enable prompts and run it from a terminal session.';
        }

        if (args.options.startTime) {
          if (!validation.isValidISODateTime(args.options.startTime)) {
            return `'${args.options.startTime}' is not a valid ISO date string for option startTime.`;
          }
          if (new Date(args.options.startTime) > new Date()) {
            return 'startTime value cannot be in the future.';
          }
        }

        if (args.options.endTime) {
          if (!validation.isValidISODateTime(args.options.endTime)) {
            return `'${args.options.endTime}' is not a valid ISO date string for option endTime.`;
          }
          if (new Date(args.options.endTime) > new Date()) {
            return 'endTime value cannot be in the future.';
          }
        }

        if (args.options.startTime &&
          args.options.endTime &&
          new Date(args.options.startTime) >= new Date(args.options.endTime)) {
          return 'startTime must be before endTime.';
        }

        if (typeof args.options.top !== 'undefined' &&
          (!Number.isInteger(args.options.top) || args.options.top <= 0)) {
          return 'top must be a positive integer.';
        }

        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid GUID for option userId.`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userName} is not a valid UPN for option userName.`;
        }

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('folderName', 'folderId', 'startTime', 'endTime', 'userId', 'userName');
  }

  #initOptionSets(): void {
    this.optionSets.push(
      {
        options: ['folderId', 'folderName'],
        runsWhen: (args) => args.options.folderId || args.options.folderName
      },
      {
        options: ['userId', 'userName'],
        runsWhen: (args) => args.options.userId || args.options.userName
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (!args.options.userId &&
        !args.options.userName &&
        accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken)) {
        throw 'You must specify either the userId or userName option when using app-only permissions.';
      }

      const userUrl = args.options.userId || args.options.userName
        ? `users/${args.options.userId || formatting.encodeQueryParameter(args.options.userName!)}`
        : 'me';

      const folderId = await this.getFolderId(userUrl, args.options);
      const folderUrl = folderId ? `/mailFolders/${folderId}` : '';
      const top = args.options.top || 25;
      let requestUrl = `${this.resource}/v1.0/${userUrl}${folderUrl}/messages?$top=${top}&$orderby=receivedDateTime desc`;

      if (args.options.startTime || args.options.endTime) {
        const filters: string[] = [];

        if (args.options.startTime) {
          filters.push(`receivedDateTime ge ${args.options.startTime}`);
        }
        if (args.options.endTime) {
          filters.push(`receivedDateTime lt ${args.options.endTime}`);
        }

        if (filters.length > 0) {
          requestUrl += `&$filter=${filters.join(' and ')}`;
        }
      }

      const requestOptions: CliRequestOptions = {
        url: requestUrl,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const response = await request.get<{ value: Message[]; }>(requestOptions);
      const allMessages = response.value || [];
      const skippedMessageKeys = new Set<string>();
      let blockedMessages = this.getBlockedMessages(allMessages, skippedMessageKeys);
      let domainsAdded = 0;
      let sendersAdded = 0;
      let skipped = 0;

      if (blockedMessages.length === 0) {
        await logger.log({
          domainsAdded,
          filePath: mailSenderWhitelist.getFilePath(),
          remaining: blockedMessages.length,
          sendersAdded,
          skipped
        });
        return;
      }

      while (blockedMessages.length > 0) {
        const selectedMessage = blockedMessages[0];
        const senderAddress = mailSenderWhitelist.getSenderAddress(selectedMessage);
        const senderDomain = mailSenderWhitelist.getSenderDomain(selectedMessage);

        await logger.logToStderr(`Sender: ${senderAddress || 'unknown sender'}`);
        await logger.logToStderr(`Subject: ${selectedMessage.subject || '(no subject)'}`);
        if (senderDomain) {
          await logger.logToStderr(`Whitelist target: ${senderDomain}`);
        }

        const action = await cli.promptForInput({
          message: 'Choose action: [a]dd to whitelist, add specific [s]ender to whitelist, [d]o not add to whitelist',
          validate: (value) => {
            const normalizedValue = value.trim().toLowerCase();
            if (['a', 's', 'd'].includes(normalizedValue)) {
              return true;
            }

            return 'Choose a, s, or d.';
          }
        }) as ReviewAction;

        if (action === 'd') {
          skipped++;
          skippedMessageKeys.add(this.getMessageKey(selectedMessage));
        }
        else if (action === 'a' && senderDomain) {
          if (mailSenderWhitelist.addDomain(senderDomain)) {
            domainsAdded++;
            await logger.logToStderr(`Whitelisted domain ${senderDomain}`);
          }
        }
        else if (action === 's' && senderAddress) {
          if (mailSenderWhitelist.addSender(senderAddress)) {
            sendersAdded++;
            await logger.logToStderr(`Whitelisted sender ${senderAddress}`);
          }
        }

        blockedMessages = this.getBlockedMessages(allMessages, skippedMessageKeys);
      }

      await logger.log({
        domainsAdded,
        filePath: mailSenderWhitelist.getFilePath(),
        remaining: blockedMessages.length,
        sendersAdded,
        skipped
      });
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getBlockedMessages(messages: Message[], skippedMessageKeys: Set<string>): Message[] {
    return mailSenderWhitelist
      .filterMessages(messages)
      .blocked
      .filter(message => !skippedMessageKeys.has(this.getMessageKey(message)));
  }

  private getMessageKey(message: Message): string {
    return message.id || `${mailSenderWhitelist.getSenderAddress(message) || 'unknown'}|${message.receivedDateTime || ''}|${message.subject || ''}`;
  }

  private async getFolderId(userUrl: string, options: Options): Promise<string> {
    if (!options.folderId && !options.folderName) {
      return '';
    }

    if (options.folderId) {
      return options.folderId;
    }

    if (Outlook.wellKnownFolderNames.includes(options.folderName!.toLowerCase())) {
      return options.folderName!;
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/${userUrl}/mailFolders?$filter=displayName eq '${formatting.encodeQueryParameter(options.folderName!)}'&$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: { id: string; }[] }>(requestOptions);

    if (response.value.length === 0) {
      throw `Folder with name '${options.folderName as string}' not found`;
    }

    if (response.value.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', response.value);
      const result = await cli.handleMultipleResultsFound<{ id: string; }>(`Multiple folders with name '${options.folderName!}' found.`, resultAsKeyValuePair);
      return result.id;
    }

    return response.value[0].id;
  }
}

export default new OutlookMessageWhitelistCommand();
