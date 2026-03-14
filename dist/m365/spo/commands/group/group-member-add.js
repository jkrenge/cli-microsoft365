var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoGroupMemberAddCommand_instances, _SpoGroupMemberAddCommand_initTelemetry, _SpoGroupMemberAddCommand_initOptions, _SpoGroupMemberAddCommand_initValidators, _SpoGroupMemberAddCommand_initOptionSets, _SpoGroupMemberAddCommand_initTypes;
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoGroupMemberAddCommand extends SpoCommand {
    get name() {
        return commands.GROUP_MEMBER_ADD;
    }
    get description() {
        return 'Add members to a SharePoint Group';
    }
    constructor() {
        super();
        _SpoGroupMemberAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoGroupMemberAddCommand_instances, "m", _SpoGroupMemberAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoGroupMemberAddCommand_instances, "m", _SpoGroupMemberAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoGroupMemberAddCommand_instances, "m", _SpoGroupMemberAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoGroupMemberAddCommand_instances, "m", _SpoGroupMemberAddCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoGroupMemberAddCommand_instances, "m", _SpoGroupMemberAddCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            const loginNames = await this.getLoginNames(logger, args.options);
            let apiUrl = `${args.options.webUrl}/_api/web/SiteGroups`;
            if (args.options.groupId) {
                apiUrl += `/GetById(${args.options.groupId})`;
            }
            else {
                apiUrl += `/GetByName('${formatting.encodeQueryParameter(args.options.groupName)}')`;
            }
            apiUrl += '/users';
            if (this.verbose) {
                await logger.logToStderr('Adding members to group...');
            }
            const result = [];
            for (const loginName of loginNames) {
                const requestOptions = {
                    url: apiUrl,
                    headers: {
                        accept: 'application/json;odata=nometadata'
                    },
                    responseType: 'json',
                    data: {
                        LoginName: loginName
                    }
                };
                const response = await request.post(requestOptions);
                result.push(response);
            }
            await logger.log(result);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getLoginNames(logger, options) {
        const loginNames = [];
        if (options.userNames || options.emails) {
            loginNames.push(...formatting.splitAndTrim(options.userNames || options.emails).map(u => `i:0#.f|membership|${u}`));
        }
        else if (options.entraGroupIds || options.entraGroupNames) {
            if (this.verbose) {
                await logger.logToStderr(`Resolving ${(options.entraGroupIds || options.entraGroupNames).length} group(s)...`);
            }
            const groups = [];
            if (options.entraGroupIds) {
                const groupIds = formatting.splitAndTrim(options.entraGroupIds);
                for (const groupId of groupIds) {
                    const group = await entraGroup.getGroupById(groupId);
                    groups.push(group);
                }
            }
            else {
                const groupNames = formatting.splitAndTrim(options.entraGroupNames);
                for (const groupName of groupNames) {
                    const group = await entraGroup.getGroupByDisplayName(groupName);
                    groups.push(group);
                }
            }
            // Check if group is M365 group or security group
            loginNames.push(...groups.map(g => g.mailEnabled ? `c:0o.c|federateddirectoryclaimprovider|${g.id}` : `c:0t.c|tenant|${g.id}`));
        }
        else if (options.userIds) {
            const userIds = formatting.splitAndTrim(options.userIds);
            if (this.verbose) {
                await logger.logToStderr(`Resolving ${userIds.length} user(s)...`);
            }
            for (const userId of userIds) {
                const loginName = await this.getUserLoginNameById(options.webUrl, parseInt(userId));
                loginNames.push(loginName);
            }
        }
        return loginNames;
    }
    async getUserLoginNameById(webUrl, userId) {
        const requestOptions = {
            url: `${webUrl}/_api/web/SiteUsers/GetById(${userId})?$select=LoginName`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const user = await request.get(requestOptions);
        return user.LoginName;
    }
}
_SpoGroupMemberAddCommand_instances = new WeakSet(), _SpoGroupMemberAddCommand_initTelemetry = function _SpoGroupMemberAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            groupId: typeof args.options.groupId !== 'undefined',
            groupName: typeof args.options.groupName !== 'undefined',
            userNames: typeof args.options.userNames !== 'undefined',
            emails: typeof args.options.emails !== 'undefined',
            userIds: typeof args.options.userIds !== 'undefined',
            entraGroupIds: typeof args.options.entraGroupIds !== 'undefined',
            entraGroupNames: typeof args.options.entraGroupNames !== 'undefined'
        });
    });
}, _SpoGroupMemberAddCommand_initOptions = function _SpoGroupMemberAddCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--groupId [groupId]'
    }, {
        option: '--groupName [groupName]'
    }, {
        option: '--userNames [userNames]'
    }, {
        option: '--emails [emails]'
    }, {
        option: '--userIds [userIds]'
    }, {
        option: '--entraGroupIds [entraGroupIds]'
    }, {
        option: '--entraGroupNames [entraGroupNames]'
    });
}, _SpoGroupMemberAddCommand_initValidators = function _SpoGroupMemberAddCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.groupId && !validation.isValidPositiveInteger(args.options.groupId)) {
            return `Specified groupId ${args.options.groupId} is not a positive number.`;
        }
        if (args.options.userIds) {
            const isValidArray = validation.isValidPositiveIntegerArray(args.options.userIds);
            if (isValidArray !== true) {
                return `Option 'userIds' contains one or more invalid numbers: ${isValidArray}.`;
            }
        }
        if (args.options.userNames) {
            const isValidArray = validation.isValidUserPrincipalNameArray(args.options.userNames);
            if (isValidArray !== true) {
                return `Option 'userNames' contains one or more invalid UPNs: ${isValidArray}.`;
            }
        }
        if (args.options.emails) {
            const isValidArray = validation.isValidUserPrincipalNameArray(args.options.emails);
            if (isValidArray !== true) {
                return `Option 'emails' contains one or more invalid UPNs: ${isValidArray}.`;
            }
        }
        if (args.options.entraGroupIds) {
            const isValidArray = validation.isValidGuidArray(args.options.entraGroupIds);
            if (isValidArray !== true) {
                return `Option 'entraGroupIds' contains one or more invalid GUIDs: ${isValidArray}.`;
            }
        }
        return true;
    });
}, _SpoGroupMemberAddCommand_initOptionSets = function _SpoGroupMemberAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['groupId', 'groupName'] }, { options: ['userNames', 'emails', 'userIds', 'entraGroupIds', 'entraGroupNames'] });
}, _SpoGroupMemberAddCommand_initTypes = function _SpoGroupMemberAddCommand_initTypes() {
    this.types.string.push('webUrl', 'groupName', 'userNames', 'emails', 'userIds', 'entraGroupIds', 'entraGroupNames');
};
export default new SpoGroupMemberAddCommand();
//# sourceMappingURL=group-member-add.js.map