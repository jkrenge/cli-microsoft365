var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _OneNoteNotebookAddCommand_instances, _OneNoteNotebookAddCommand_initTelemetry, _OneNoteNotebookAddCommand_initOptions, _OneNoteNotebookAddCommand_initValidators, _OneNoteNotebookAddCommand_initOptionSets;
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import { spo } from '../../../../utils/spo.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import { formatting } from '../../../../utils/formatting.js';
class OneNoteNotebookAddCommand extends GraphDelegatedCommand {
    get name() {
        return commands.NOTEBOOK_ADD;
    }
    get description() {
        return 'Create a new OneNote notebook';
    }
    constructor() {
        super();
        _OneNoteNotebookAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _OneNoteNotebookAddCommand_instances, "m", _OneNoteNotebookAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _OneNoteNotebookAddCommand_instances, "m", _OneNoteNotebookAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _OneNoteNotebookAddCommand_instances, "m", _OneNoteNotebookAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _OneNoteNotebookAddCommand_instances, "m", _OneNoteNotebookAddCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Creating OneNote notebook ${args.options.name}`);
            }
            const requestUrl = await this.getRequestUrl(args);
            const requestOptions = {
                url: requestUrl,
                headers: {
                    accept: 'application/json;odata.metadata=none',
                    'content-type': "application/json"
                },
                responseType: 'json',
                data: {
                    displayName: args.options.name
                }
            };
            const response = await request.post(requestOptions);
            await logger.log(response);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getRequestUrl(args) {
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
        endpoint += '/onenote/notebooks';
        return endpoint;
    }
}
_OneNoteNotebookAddCommand_instances = new WeakSet(), _OneNoteNotebookAddCommand_initTelemetry = function _OneNoteNotebookAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            groupId: typeof args.options.groupId !== 'undefined',
            groupName: typeof args.options.groupName !== 'undefined',
            webUrl: typeof args.options.webUrl !== 'undefined'
        });
    });
}, _OneNoteNotebookAddCommand_initOptions = function _OneNoteNotebookAddCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name <name>'
    }, {
        option: '--userId [userId]'
    }, {
        option: '--userName [userName]'
    }, {
        option: '--groupId [groupId]'
    }, {
        option: '--groupName [groupName]'
    }, {
        option: '-u, --webUrl [webUrl]'
    });
}, _OneNoteNotebookAddCommand_initValidators = function _OneNoteNotebookAddCommand_initValidators() {
    this.validators.push(async (args) => {
        // check name for invalid characters
        if (args.options.name.length > 128) {
            return 'The specified name is too long. It should be less than 128 characters';
        }
        if (/[?*/:<>|'"]/.test(args.options.name)) {
            return `The specified name contains invalid characters. It cannot contain ?*/:<>|'". Please remove them and try again.`;
        }
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
}, _OneNoteNotebookAddCommand_initOptionSets = function _OneNoteNotebookAddCommand_initOptionSets() {
    this.optionSets.push({
        options: ['userId', 'userName', 'groupId', 'groupName', 'webUrl'],
        runsWhen: (args) => {
            const options = [args.options.userId, args.options.userName, args.options.groupId, args.options.groupName, args.options.webUrl];
            return options.some(item => item !== undefined);
        }
    });
};
export default new OneNoteNotebookAddCommand();
//# sourceMappingURL=notebook-add.js.map