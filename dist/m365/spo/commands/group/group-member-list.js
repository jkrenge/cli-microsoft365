var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoGroupMemberListCommand_instances, _SpoGroupMemberListCommand_initTelemetry, _SpoGroupMemberListCommand_initOptions, _SpoGroupMemberListCommand_initValidators, _SpoGroupMemberListCommand_initOptionSets;
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoGroupMemberListCommand extends SpoCommand {
    get name() {
        return commands.GROUP_MEMBER_LIST;
    }
    get description() {
        return `List the members of a SharePoint Group`;
    }
    defaultProperties() {
        return ['Title', 'UserPrincipalName', 'Id', 'Email'];
    }
    constructor() {
        super();
        _SpoGroupMemberListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoGroupMemberListCommand_instances, "m", _SpoGroupMemberListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoGroupMemberListCommand_instances, "m", _SpoGroupMemberListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoGroupMemberListCommand_instances, "m", _SpoGroupMemberListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoGroupMemberListCommand_instances, "m", _SpoGroupMemberListCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving the list of members from the SharePoint group :  ${args.options.groupId ? args.options.groupId : args.options.groupName}`);
        }
        const requestUrl = `${args.options.webUrl}/_api/web/sitegroups/${args.options.groupId
            ? `GetById('${args.options.groupId}')`
            : `GetByName('${formatting.encodeQueryParameter(args.options.groupName)}')`}/users`;
        try {
            const response = await odata.getAllItems(requestUrl);
            await logger.log(response);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoGroupMemberListCommand_instances = new WeakSet(), _SpoGroupMemberListCommand_initTelemetry = function _SpoGroupMemberListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            groupId: typeof args.options.groupId !== 'undefined',
            groupName: typeof args.options.groupName !== 'undefined'
        });
    });
}, _SpoGroupMemberListCommand_initOptions = function _SpoGroupMemberListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--groupId [groupId]'
    }, {
        option: '--groupName [groupName]'
    });
}, _SpoGroupMemberListCommand_initValidators = function _SpoGroupMemberListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.groupId && isNaN(args.options.groupId)) {
            return `Specified "groupId" ${args.options.groupId} is not valid`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
}, _SpoGroupMemberListCommand_initOptionSets = function _SpoGroupMemberListCommand_initOptionSets() {
    this.optionSets.push({ options: ['groupName', 'groupId'] });
};
export default new SpoGroupMemberListCommand();
//# sourceMappingURL=group-member-list.js.map