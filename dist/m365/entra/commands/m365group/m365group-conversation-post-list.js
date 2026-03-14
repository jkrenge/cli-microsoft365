var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraM365GroupConversationPostListCommand_instances, _EntraM365GroupConversationPostListCommand_initTelemetry, _EntraM365GroupConversationPostListCommand_initOptions, _EntraM365GroupConversationPostListCommand_initValidators, _EntraM365GroupConversationPostListCommand_initOptionSets;
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraM365GroupConversationPostListCommand extends GraphCommand {
    get name() {
        return commands.M365GROUP_CONVERSATION_POST_LIST;
    }
    get description() {
        return 'Lists conversation posts of a Microsoft 365 group';
    }
    constructor() {
        super();
        _EntraM365GroupConversationPostListCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraM365GroupConversationPostListCommand_instances, "m", _EntraM365GroupConversationPostListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupConversationPostListCommand_instances, "m", _EntraM365GroupConversationPostListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupConversationPostListCommand_instances, "m", _EntraM365GroupConversationPostListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupConversationPostListCommand_instances, "m", _EntraM365GroupConversationPostListCommand_initOptionSets).call(this);
    }
    defaultProperties() {
        return ['receivedDateTime', 'id'];
    }
    async commandAction(logger, args) {
        try {
            const retrievedgroupId = await this.getGroupId(args);
            const isUnifiedGroup = await entraGroup.isUnifiedGroup(retrievedgroupId);
            if (!isUnifiedGroup) {
                throw Error(`Specified group with id '${retrievedgroupId}' is not a Microsoft 365 group.`);
            }
            const posts = await odata.getAllItems(`${this.resource}/v1.0/groups/${retrievedgroupId}/threads/${args.options.threadId}/posts`);
            await logger.log(posts);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getGroupId(args) {
        if (args.options.groupId) {
            return formatting.encodeQueryParameter(args.options.groupId);
        }
        const group = await entraGroup.getGroupByDisplayName(args.options.groupName);
        return group.id;
    }
}
_EntraM365GroupConversationPostListCommand_instances = new WeakSet(), _EntraM365GroupConversationPostListCommand_initTelemetry = function _EntraM365GroupConversationPostListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            groupId: typeof args.options.groupId !== 'undefined',
            groupName: typeof args.options.groupName !== 'undefined'
        });
    });
}, _EntraM365GroupConversationPostListCommand_initOptions = function _EntraM365GroupConversationPostListCommand_initOptions() {
    this.options.unshift({
        option: '-i, --groupId [groupId]'
    }, {
        option: '-d, --groupName [groupName]'
    }, {
        option: '-t, --threadId <threadId>'
    });
}, _EntraM365GroupConversationPostListCommand_initValidators = function _EntraM365GroupConversationPostListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.groupId && !validation.isValidGuid(args.options.groupId)) {
            return `${args.options.groupId} is not a valid GUID`;
        }
        return true;
    });
}, _EntraM365GroupConversationPostListCommand_initOptionSets = function _EntraM365GroupConversationPostListCommand_initOptionSets() {
    this.optionSets.push({ options: ['groupId', 'groupName'] });
};
export default new EntraM365GroupConversationPostListCommand();
//# sourceMappingURL=m365group-conversation-post-list.js.map