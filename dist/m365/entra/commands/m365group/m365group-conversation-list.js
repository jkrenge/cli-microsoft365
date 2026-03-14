var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraM365GroupConversationListCommand_instances, _EntraM365GroupConversationListCommand_initOptions, _EntraM365GroupConversationListCommand_initValidators, _EntraM365GroupConversationListCommand_initOptionSets, _EntraM365GroupConversationListCommand_initTypes;
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
class EntraM365GroupConversationListCommand extends GraphCommand {
    get name() {
        return commands.M365GROUP_CONVERSATION_LIST;
    }
    get description() {
        return 'Lists conversations for the specified Microsoft 365 group';
    }
    defaultProperties() {
        return ['topic', 'lastDeliveredDateTime', 'id'];
    }
    constructor() {
        super();
        _EntraM365GroupConversationListCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraM365GroupConversationListCommand_instances, "m", _EntraM365GroupConversationListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupConversationListCommand_instances, "m", _EntraM365GroupConversationListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupConversationListCommand_instances, "m", _EntraM365GroupConversationListCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupConversationListCommand_instances, "m", _EntraM365GroupConversationListCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving conversations for Microsoft 365 Group: ${args.options.groupId || args.options.groupName}...`);
        }
        try {
            let groupId = args.options.groupId;
            if (args.options.groupName) {
                groupId = await entraGroup.getGroupIdByDisplayName(args.options.groupName);
            }
            const isUnifiedGroup = await entraGroup.isUnifiedGroup(groupId);
            if (!isUnifiedGroup) {
                throw Error(`Specified group with id '${groupId}' is not a Microsoft 365 group.`);
            }
            const conversations = await odata.getAllItems(`${this.resource}/v1.0/groups/${groupId}/conversations`);
            await logger.log(conversations);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_EntraM365GroupConversationListCommand_instances = new WeakSet(), _EntraM365GroupConversationListCommand_initOptions = function _EntraM365GroupConversationListCommand_initOptions() {
    this.options.unshift({
        option: '-i, --groupId [groupId]'
    }, {
        option: '-n, --groupName [groupName]'
    });
}, _EntraM365GroupConversationListCommand_initValidators = function _EntraM365GroupConversationListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.groupId && !validation.isValidGuid(args.options.groupId)) {
            return `${args.options.groupId} is not a valid GUID`;
        }
        return true;
    });
}, _EntraM365GroupConversationListCommand_initOptionSets = function _EntraM365GroupConversationListCommand_initOptionSets() {
    this.optionSets.push({ options: ['groupId', 'groupName'] });
}, _EntraM365GroupConversationListCommand_initTypes = function _EntraM365GroupConversationListCommand_initTypes() {
    this.types.string.push('groupId', 'groupName');
};
export default new EntraM365GroupConversationListCommand();
//# sourceMappingURL=m365group-conversation-list.js.map