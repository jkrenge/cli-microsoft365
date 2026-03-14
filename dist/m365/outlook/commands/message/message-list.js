var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _OutlookMessageListCommand_instances, _OutlookMessageListCommand_initTelemetry, _OutlookMessageListCommand_initOptions, _OutlookMessageListCommand_initValidators, _OutlookMessageListCommand_initTypes, _OutlookMessageListCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { Outlook } from '../../Outlook.js';
import { cli } from '../../../../cli/cli.js';
import { validation } from '../../../../utils/validation.js';
import { accessToken } from '../../../../utils/accessToken.js';
import auth from '../../../../Auth.js';
import { mailSenderWhitelist } from '../../../../utils/mailSenderWhitelist.js';
class OutlookMessageListCommand extends GraphCommand {
    get name() {
        return commands.MESSAGE_LIST;
    }
    get description() {
        return 'Gets all mail messages from the specified folder';
    }
    constructor() {
        super();
        _OutlookMessageListCommand_instances.add(this);
        __classPrivateFieldGet(this, _OutlookMessageListCommand_instances, "m", _OutlookMessageListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _OutlookMessageListCommand_instances, "m", _OutlookMessageListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _OutlookMessageListCommand_instances, "m", _OutlookMessageListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _OutlookMessageListCommand_instances, "m", _OutlookMessageListCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _OutlookMessageListCommand_instances, "m", _OutlookMessageListCommand_initOptionSets).call(this);
    }
    defaultProperties() {
        return ['subject', 'receivedDateTime'];
    }
    async commandAction(logger, args) {
        try {
            if (!args.options.userId && !args.options.userName && accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken)) {
                throw 'You must specify either the userId or userName option when using app-only permissions.';
            }
            const userUrl = args.options.userId || args.options.userName ? `users/${args.options.userId || formatting.encodeQueryParameter(args.options.userName)}` : 'me';
            const folderId = await this.getFolderId(userUrl, args.options);
            const folderUrl = folderId ? `/mailFolders/${folderId}` : '';
            let requestUrl = `${this.resource}/v1.0/${userUrl}${folderUrl}/messages?$top=100`;
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
            let messages = await odata.getAllItems(requestUrl);
            if (mailSenderWhitelist.shouldFilterInInteractiveMode()) {
                const filterResult = mailSenderWhitelist.filterMessages(messages);
                messages = filterResult.allowed;
                await mailSenderWhitelist.logFilteredSummary(logger, filterResult.filteredCount);
            }
            await logger.log(messages);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
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
_OutlookMessageListCommand_instances = new WeakSet(), _OutlookMessageListCommand_initTelemetry = function _OutlookMessageListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            folderId: typeof args.options.folderId !== 'undefined',
            folderName: typeof args.options.folderName !== 'undefined',
            startTime: typeof args.options.startTime !== 'undefined',
            endTime: typeof args.options.endTime !== 'undefined',
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined'
        });
    });
}, _OutlookMessageListCommand_initOptions = function _OutlookMessageListCommand_initOptions() {
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
        option: '--userId [userId]'
    }, {
        option: '--userName [userName]'
    });
}, _OutlookMessageListCommand_initValidators = function _OutlookMessageListCommand_initValidators() {
    this.validators.push(async (args) => {
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
        if (args.options.startTime && args.options.endTime && new Date(args.options.startTime) >= new Date(args.options.endTime)) {
            return 'startTime must be before endTime.';
        }
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID for option userId.`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid UPN for option userName.`;
        }
        return true;
    });
}, _OutlookMessageListCommand_initTypes = function _OutlookMessageListCommand_initTypes() {
    this.types.string.push('folderName', 'folderId', 'startTime', 'endTime', 'userId', 'userName');
}, _OutlookMessageListCommand_initOptionSets = function _OutlookMessageListCommand_initOptionSets() {
    this.optionSets.push({
        options: ['folderId', 'folderName'],
        runsWhen: (args) => args.options.folderId || args.options.folderName
    }, {
        options: ['userId', 'userName'],
        runsWhen: (args) => args.options.userId || args.options.userName
    });
};
export default new OutlookMessageListCommand();
//# sourceMappingURL=message-list.js.map