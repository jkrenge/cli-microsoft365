var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TodoTaskGetCommand_instances, _TodoTaskGetCommand_initTelemetry, _TodoTaskGetCommand_initOptions, _TodoTaskGetCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import commands from '../../commands.js';
class TodoTaskGetCommand extends GraphDelegatedCommand {
    get name() {
        return commands.TASK_GET;
    }
    get description() {
        return 'Get a specific task from a Microsoft To Do task list';
    }
    constructor() {
        super();
        _TodoTaskGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _TodoTaskGetCommand_instances, "m", _TodoTaskGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TodoTaskGetCommand_instances, "m", _TodoTaskGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TodoTaskGetCommand_instances, "m", _TodoTaskGetCommand_initOptionSets).call(this);
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
            const requestOptions = {
                url: `${this.resource}/v1.0/me/todo/lists/${listId}/tasks/${args.options.id}`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            const item = await request.get(requestOptions);
            if (!cli.shouldTrimOutput(args.options.output)) {
                await logger.log(item);
            }
            else {
                await logger.log({
                    id: item.id,
                    title: item.title,
                    status: item.status,
                    createdDateTime: item.createdDateTime,
                    lastModifiedDateTime: item.lastModifiedDateTime
                });
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_TodoTaskGetCommand_instances = new WeakSet(), _TodoTaskGetCommand_initTelemetry = function _TodoTaskGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listName: typeof args.options.listName !== 'undefined'
        });
    });
}, _TodoTaskGetCommand_initOptions = function _TodoTaskGetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '--listName [listName]'
    }, {
        option: '--listId [listId]'
    });
}, _TodoTaskGetCommand_initOptionSets = function _TodoTaskGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listName'] });
};
export default new TodoTaskGetCommand();
//# sourceMappingURL=task-get.js.map