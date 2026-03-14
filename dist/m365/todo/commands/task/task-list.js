var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TodoTaskListCommand_instances, _TodoTaskListCommand_initTelemetry, _TodoTaskListCommand_initOptions, _TodoTaskListCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import commands from '../../commands.js';
class TodoTaskListCommand extends GraphDelegatedCommand {
    get name() {
        return commands.TASK_LIST;
    }
    get description() {
        return 'List tasks from a Microsoft To Do task list';
    }
    defaultProperties() {
        return ['id', 'title', 'status', 'createdDateTime', 'lastModifiedDateTime'];
    }
    constructor() {
        super();
        _TodoTaskListCommand_instances.add(this);
        __classPrivateFieldGet(this, _TodoTaskListCommand_instances, "m", _TodoTaskListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TodoTaskListCommand_instances, "m", _TodoTaskListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TodoTaskListCommand_instances, "m", _TodoTaskListCommand_initOptionSets).call(this);
    }
    async getTodoListId(args) {
        if (args.options.listId) {
            return args.options.listId;
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/me/todo/lists?$filter=displayName eq '${formatting.encodeQueryParameter(args.options.listName)}'`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        const taskList = response.value[0];
        if (!taskList) {
            throw `The specified task list does not exist`;
        }
        return taskList.id;
    }
    async commandAction(logger, args) {
        try {
            const listId = await this.getTodoListId(args);
            const endpoint = `${this.resource}/v1.0/me/todo/lists/${listId}/tasks`;
            const items = await odata.getAllItems(endpoint);
            if (!cli.shouldTrimOutput(args.options.output)) {
                await logger.log(items);
            }
            else {
                await logger.log(items.map(m => {
                    return {
                        id: m.id,
                        title: m.title,
                        status: m.status,
                        createdDateTime: m.createdDateTime,
                        lastModifiedDateTime: m.lastModifiedDateTime
                    };
                }));
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_TodoTaskListCommand_instances = new WeakSet(), _TodoTaskListCommand_initTelemetry = function _TodoTaskListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listName: typeof args.options.listName !== 'undefined'
        });
    });
}, _TodoTaskListCommand_initOptions = function _TodoTaskListCommand_initOptions() {
    this.options.unshift({
        option: '--listName [listName]'
    }, {
        option: '--listId [listId]'
    });
}, _TodoTaskListCommand_initOptionSets = function _TodoTaskListCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listName'] });
};
export default new TodoTaskListCommand();
//# sourceMappingURL=task-list.js.map