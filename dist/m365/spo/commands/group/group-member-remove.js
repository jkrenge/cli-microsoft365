var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoGroupMemberRemoveCommand_instances, _SpoGroupMemberRemoveCommand_initTelemetry, _SpoGroupMemberRemoveCommand_initOptions, _SpoGroupMemberRemoveCommand_initValidators, _SpoGroupMemberRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import entraUserGetCommand from '../../../entra/commands/user/user-get.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import SpoGroupMemberListCommand from './group-member-list.js';
class SpoGroupMemberRemoveCommand extends SpoCommand {
    get name() {
        return commands.GROUP_MEMBER_REMOVE;
    }
    get description() {
        return 'Removes the specified member from a SharePoint group';
    }
    constructor() {
        super();
        _SpoGroupMemberRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoGroupMemberRemoveCommand_instances, "m", _SpoGroupMemberRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoGroupMemberRemoveCommand_instances, "m", _SpoGroupMemberRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoGroupMemberRemoveCommand_instances, "m", _SpoGroupMemberRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoGroupMemberRemoveCommand_instances, "m", _SpoGroupMemberRemoveCommand_initOptionSets).call(this);
    }
    async getUserName(logger, args) {
        if (args.options.userName) {
            return args.options.userName;
        }
        if (this.verbose) {
            await logger.logToStderr(`Retrieving information about the user ${args.options.email}`);
        }
        const options = {
            email: args.options.email,
            output: 'json',
            debug: args.options.debug,
            verbose: args.options.verbose
        };
        const userGetOutput = await cli.executeCommandWithOutput(entraUserGetCommand, { options: { ...options, _: [] } });
        const userOutput = JSON.parse(userGetOutput.stdout);
        return userOutput.userPrincipalName;
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            if (this.debug) {
                await logger.logToStderr('Confirmation bypassed by entering confirm option. Removing the user from SharePoint Group...');
            }
            await this.removeUserfromSPGroup(logger, args);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove user ${args.options.userName || args.options.userId || args.options.email || args.options.entraGroupId || args.options.entraGroupName} from the SharePoint group?` });
            if (result) {
                await this.removeUserfromSPGroup(logger, args);
            }
        }
    }
    async removeUserfromSPGroup(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Removing User ${args.options.userName || args.options.email || args.options.userId || args.options.entraGroupId || args.options.entraGroupName} from Group: ${args.options.groupId || args.options.groupName}`);
        }
        let requestUrl = `${args.options.webUrl}/_api/web/sitegroups/${args.options.groupId
            ? `GetById('${args.options.groupId}')`
            : `GetByName('${formatting.encodeQueryParameter(args.options.groupName)}')`}`;
        if (args.options.userId) {
            requestUrl += `/users/removeById(${args.options.userId})`;
        }
        else if (args.options.userName || args.options.email) {
            const userName = await this.getUserName(logger, args);
            const loginName = `i:0#.f|membership|${userName}`;
            requestUrl += `/users/removeByLoginName(@LoginName)?@LoginName='${formatting.encodeQueryParameter(loginName)}'`;
        }
        else {
            const entraGroupId = await this.getGroupId(args);
            requestUrl += `/users/RemoveById(${entraGroupId})`;
        }
        const requestOptions = {
            url: requestUrl,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        try {
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getGroupId(args) {
        const options = {
            webUrl: args.options.webUrl,
            output: 'json',
            debug: this.debug,
            verbose: this.verbose
        };
        if (args.options.groupId) {
            options.groupId = args.options.groupId;
        }
        else {
            options.groupName = args.options.groupName;
        }
        const output = await cli.executeCommandWithOutput(SpoGroupMemberListCommand, { options: { ...options, _: [] } });
        const getGroupMemberListOutput = JSON.parse(output.stdout);
        let foundGroups;
        if (args.options.entraGroupId) {
            foundGroups = getGroupMemberListOutput.filter((x) => { return x.LoginName.indexOf(args.options.entraGroupId) > -1 && (x.LoginName.indexOf("c:0o.c|federateddirectoryclaimprovider|") === 0 || x.LoginName.indexOf("c:0t.c|tenant|") === 0); });
        }
        else {
            foundGroups = getGroupMemberListOutput.filter((x) => { return x.Title === args.options.entraGroupName && (x.LoginName.indexOf("c:0o.c|federateddirectoryclaimprovider|") === 0 || x.LoginName.indexOf("c:0t.c|tenant|") === 0); });
        }
        if (foundGroups.length === 0) {
            throw `The Microsoft Entra group ${args.options.entraGroupId || args.options.entraGroupName} is not found in SharePoint group ${args.options.groupId || args.options.groupName}`;
        }
        return foundGroups[0].Id;
    }
}
_SpoGroupMemberRemoveCommand_instances = new WeakSet(), _SpoGroupMemberRemoveCommand_initTelemetry = function _SpoGroupMemberRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            groupId: (!(!args.options.groupId)).toString(),
            groupName: (!(!args.options.groupName)).toString(),
            userName: (!(!args.options.userName)).toString(),
            email: (!(!args.options.email)).toString(),
            userId: (!(!args.options.userId)).toString(),
            entraGroupId: (!(!args.options.entraGroupId)).toString(),
            entraGroupName: (!(!args.options.entraGroupName)).toString(),
            force: (!(!args.options.force)).toString()
        });
    });
}, _SpoGroupMemberRemoveCommand_initOptions = function _SpoGroupMemberRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--groupId [groupId]'
    }, {
        option: '--groupName [groupName]'
    }, {
        option: '--userName [userName]'
    }, {
        option: '--email [email]'
    }, {
        option: '--userId [userId]'
    }, {
        option: '--entraGroupId [entraGroupId]'
    }, {
        option: '--entraGroupName [entraGroupName]'
    }, {
        option: '-f, --force'
    });
}, _SpoGroupMemberRemoveCommand_initValidators = function _SpoGroupMemberRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.groupId && isNaN(args.options.groupId)) {
            return `Specified "groupId" ${args.options.groupId} is not valid`;
        }
        if (args.options.userId && isNaN(args.options.userId)) {
            return `Specified "userId" ${args.options.userId} is not valid`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid userName`;
        }
        if (args.options.email && !validation.isValidUserPrincipalName(args.options.email)) {
            return `${args.options.email} is not a valid email`;
        }
        if (args.options.entraGroupId && !validation.isValidGuid(args.options.entraGroupId)) {
            return `${args.options.entraGroupId} is not a valid GUID`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
}, _SpoGroupMemberRemoveCommand_initOptionSets = function _SpoGroupMemberRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['groupName', 'groupId'] }, { options: ['userName', 'email', 'userId', 'entraGroupId', 'entraGroupName'] });
};
export default new SpoGroupMemberRemoveCommand();
//# sourceMappingURL=group-member-remove.js.map