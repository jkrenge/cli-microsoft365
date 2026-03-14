var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _OutlookMessageWhitelistCommand_instances, _OutlookMessageWhitelistCommand_initTelemetry, _OutlookMessageWhitelistCommand_initOptions, _OutlookMessageWhitelistCommand_initValidators, _OutlookMessageWhitelistCommand_initTypes, _OutlookMessageWhitelistCommand_initOptionSets;
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { mailSenderWhitelist } from '../../../../utils/mailSenderWhitelist.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { Outlook } from '../../Outlook.js';
class OutlookMessageWhitelistCommand extends GraphCommand {
    get name() {
        return commands.MESSAGE_WHITELIST;
    }
    get description() {
        return 'Interactively review and whitelist mail senders';
    }
    constructor() {
        super();
        _OutlookMessageWhitelistCommand_instances.add(this);
        __classPrivateFieldGet(this, _OutlookMessageWhitelistCommand_instances, "m", _OutlookMessageWhitelistCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _OutlookMessageWhitelistCommand_instances, "m", _OutlookMessageWhitelistCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _OutlookMessageWhitelistCommand_instances, "m", _OutlookMessageWhitelistCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _OutlookMessageWhitelistCommand_instances, "m", _OutlookMessageWhitelistCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _OutlookMessageWhitelistCommand_instances, "m", _OutlookMessageWhitelistCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (!args.options.userId &&
                !args.options.userName &&
                accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken)) {
                throw 'You must specify either the userId or userName option when using app-only permissions.';
            }
            const userUrl = args.options.userId || args.options.userName
                ? `users/${args.options.userId || formatting.encodeQueryParameter(args.options.userName)}`
                : 'me';
            const folderId = await this.getFolderId(userUrl, args.options);
            const folderUrl = folderId ? `/mailFolders/${folderId}` : '';
            const top = args.options.top || 25;
            let requestUrl = `${this.resource}/v1.0/${userUrl}${folderUrl}/messages?$top=${top}&$orderby=receivedDateTime desc`;
            if (args.options.startTime || args.options.endTime) {
                const filters = [];
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
            const requestOptions = {
                url: requestUrl,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            const response = await request.get(requestOptions);
            const allMessages = response.value || [];
            const skippedMessageKeys = new Set();
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
                });
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
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getBlockedMessages(messages, skippedMessageKeys) {
        return mailSenderWhitelist
            .filterMessages(messages)
            .blocked
            .filter(message => !skippedMessageKeys.has(this.getMessageKey(message)));
    }
    getMessageKey(message) {
        return message.id || `${mailSenderWhitelist.getSenderAddress(message) || 'unknown'}|${message.receivedDateTime || ''}|${message.subject || ''}`;
    }
    async getFolderId(userUrl, options) {
        if (!options.folderId && !options.folderName) {
            return '';
        }
        if (options.folderId) {
            return options.folderId;
        }
        if (Outlook.wellKnownFolderNames.includes(options.folderName.toLowerCase())) {
            return options.folderName;
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/${userUrl}/mailFolders?$filter=displayName eq '${formatting.encodeQueryParameter(options.folderName)}'&$select=id`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        if (response.value.length === 0) {
            throw `Folder with name '${options.folderName}' not found`;
        }
        if (response.value.length > 1) {
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', response.value);
            const result = await cli.handleMultipleResultsFound(`Multiple folders with name '${options.folderName}' found.`, resultAsKeyValuePair);
            return result.id;
        }
        return response.value[0].id;
    }
}
_OutlookMessageWhitelistCommand_instances = new WeakSet(), _OutlookMessageWhitelistCommand_initTelemetry = function _OutlookMessageWhitelistCommand_initTelemetry() {
    this.telemetry.push((args) => {
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
}, _OutlookMessageWhitelistCommand_initOptions = function _OutlookMessageWhitelistCommand_initOptions() {
    this.options.unshift({
        option: '--folderName [folderName]',
        autocomplete: Outlook.wellKnownFolderNames
    }, {
        option: '--folderId [folderId]'
    }, {
        option: '--startTime [startTime]'
    }, {
        option: '--endTime [endTime]'
    }, {
        option: '--top [top]'
    }, {
        option: '--userId [userId]'
    }, {
        option: '--userName [userName]'
    });
}, _OutlookMessageWhitelistCommand_initValidators = function _OutlookMessageWhitelistCommand_initValidators() {
    this.validators.push(async (args) => {
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
    });
}, _OutlookMessageWhitelistCommand_initTypes = function _OutlookMessageWhitelistCommand_initTypes() {
    this.types.string.push('folderName', 'folderId', 'startTime', 'endTime', 'userId', 'userName');
}, _OutlookMessageWhitelistCommand_initOptionSets = function _OutlookMessageWhitelistCommand_initOptionSets() {
    this.optionSets.push({
        options: ['folderId', 'folderName'],
        runsWhen: (args) => args.options.folderId || args.options.folderName
    }, {
        options: ['userId', 'userName'],
        runsWhen: (args) => args.options.userId || args.options.userName
    });
};
export default new OutlookMessageWhitelistCommand();
//# sourceMappingURL=message-whitelist.js.map