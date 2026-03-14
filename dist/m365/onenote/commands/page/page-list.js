var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _OneNotePageListCommand_instances, _OneNotePageListCommand_initTelemetry, _OneNotePageListCommand_initOptions, _OneNotePageListCommand_initValidators, _OneNotePageListCommand_initOptionSets;
import { entraGroup } from '../../../../utils/entraGroup.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import { formatting } from '../../../../utils/formatting.js';
class OneNotePageListCommand extends GraphDelegatedCommand {
    get name() {
        return commands.PAGE_LIST;
    }
    get description() {
        return 'Retrieve a list of OneNote pages.';
    }
    defaultProperties() {
        return ['createdDateTime', 'title', 'id'];
    }
    constructor() {
        super();
        _OneNotePageListCommand_instances.add(this);
        __classPrivateFieldGet(this, _OneNotePageListCommand_instances, "m", _OneNotePageListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _OneNotePageListCommand_instances, "m", _OneNotePageListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _OneNotePageListCommand_instances, "m", _OneNotePageListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _OneNotePageListCommand_instances, "m", _OneNotePageListCommand_initOptionSets).call(this);
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
            const groupId = await entraGroup.getGroupIdByDisplayName(args.options.groupName);
            endpoint += `groups/${groupId}`;
        }
        else if (args.options.webUrl) {
            const siteId = await spo.getSpoGraphSiteId(args.options.webUrl);
            endpoint += `sites/${siteId}`;
        }
        else {
            endpoint += 'me';
        }
        endpoint += '/onenote/pages';
        return endpoint;
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
_OneNotePageListCommand_instances = new WeakSet(), _OneNotePageListCommand_initTelemetry = function _OneNotePageListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            groupId: typeof args.options.groupId !== 'undefined',
            groupName: typeof args.options.groupName !== 'undefined',
            webUrl: typeof args.options.webUrl !== 'undefined'
        });
    });
}, _OneNotePageListCommand_initOptions = function _OneNotePageListCommand_initOptions() {
    this.options.unshift({ option: '--userId [userId]' }, { option: '--userName [userName]' }, { option: '--groupId [groupId]' }, { option: '--groupName [groupName]' }, { option: '-u, --webUrl [webUrl]' });
}, _OneNotePageListCommand_initValidators = function _OneNotePageListCommand_initValidators() {
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
}, _OneNotePageListCommand_initOptionSets = function _OneNotePageListCommand_initOptionSets() {
    this.optionSets.push({
        options: ['userId', 'userName', 'groupId', 'groupName', 'webUrl'],
        runsWhen: (args) => {
            const options = [args.options.userId, args.options.userName, args.options.groupId, args.options.groupName, args.options.webUrl];
            return options.some(item => item !== undefined);
        }
    });
};
export default new OneNotePageListCommand();
//# sourceMappingURL=page-list.js.map