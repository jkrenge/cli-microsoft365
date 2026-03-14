var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _OneNoteNotebookListCommand_instances, _OneNoteNotebookListCommand_initTelemetry, _OneNoteNotebookListCommand_initOptions, _OneNoteNotebookListCommand_initValidators, _OneNoteNotebookListCommand_initOptionSets;
import { entraGroup } from '../../../../utils/entraGroup.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import { spo } from '../../../../utils/spo.js';
class OneNoteNotebookListCommand extends GraphDelegatedCommand {
    get name() {
        return commands.NOTEBOOK_LIST;
    }
    get description() {
        return 'Retrieve a list of notebooks';
    }
    defaultProperties() {
        return ['createdDateTime', 'displayName', 'id'];
    }
    constructor() {
        super();
        _OneNoteNotebookListCommand_instances.add(this);
        __classPrivateFieldGet(this, _OneNoteNotebookListCommand_instances, "m", _OneNoteNotebookListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _OneNoteNotebookListCommand_instances, "m", _OneNoteNotebookListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _OneNoteNotebookListCommand_instances, "m", _OneNoteNotebookListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _OneNoteNotebookListCommand_instances, "m", _OneNoteNotebookListCommand_initOptionSets).call(this);
    }
    async getEndpointUrl(args) {
        let endpoint = `${this.resource}/v1.0/`;
        if (args.options.userId) {
            endpoint += `users/${args.options.userId}`;
        }
        else if (args.options.userName) {
            endpoint += `users/${formatting.encodeQueryParameter(args.options.userName)}`;
        }
        else if (args.options.groupId) {
            endpoint += `groups/${args.options.groupId}`;
        }
        else if (args.options.groupName) {
            const groupId = await this.getGroupId(args.options.groupName);
            endpoint += `groups/${groupId}`;
        }
        else if (args.options.webUrl) {
            const siteId = await spo.getSpoGraphSiteId(args.options.webUrl);
            endpoint += `sites/${siteId}`;
        }
        else {
            endpoint += 'me';
        }
        endpoint += '/onenote/notebooks';
        return endpoint;
    }
    async getGroupId(groupName) {
        const group = await entraGroup.getGroupByDisplayName(groupName);
        return group.id;
    }
    async commandAction(logger, args) {
        try {
            const endpoint = await this.getEndpointUrl(args);
            const items = await odata.getAllItems(endpoint);
            await logger.log(items);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_OneNoteNotebookListCommand_instances = new WeakSet(), _OneNoteNotebookListCommand_initTelemetry = function _OneNoteNotebookListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            groupId: typeof args.options.groupId !== 'undefined',
            groupName: typeof args.options.groupName !== 'undefined',
            webUrl: typeof args.options.webUrl !== 'undefined'
        });
    });
}, _OneNoteNotebookListCommand_initOptions = function _OneNoteNotebookListCommand_initOptions() {
    this.options.unshift({ option: '--userId [userId]' }, { option: '--userName [userName]' }, { option: '--groupId [groupId]' }, { option: '--groupName [groupName]' }, { option: '-u, --webUrl [webUrl]' });
}, _OneNoteNotebookListCommand_initValidators = function _OneNoteNotebookListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID`;
        }
        if (args.options.groupId && !validation.isValidGuid(args.options.groupId)) {
            return `${args.options.groupId} is not a valid GUID`;
        }
        if (args.options.webUrl) {
            return validation.isValidSharePointUrl(args.options.webUrl);
        }
        return true;
    });
}, _OneNoteNotebookListCommand_initOptionSets = function _OneNoteNotebookListCommand_initOptionSets() {
    this.optionSets.push({
        options: ['userId', 'userName', 'groupId', 'groupName', 'webUrl'],
        runsWhen: (args) => {
            const options = [args.options.userId, args.options.userName, args.options.groupId, args.options.groupName, args.options.webUrl];
            return options.some(item => item !== undefined);
        }
    });
};
export default new OneNoteNotebookListCommand();
//# sourceMappingURL=notebook-list.js.map