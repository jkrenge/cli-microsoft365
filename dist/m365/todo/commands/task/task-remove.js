var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TodoTaskRemoveCommand_instances, _TodoTaskRemoveCommand_initTelemetry, _TodoTaskRemoveCommand_initOptions, _TodoTaskRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import commands from '../../commands.js';
class TodoTaskRemoveCommand extends GraphDelegatedCommand {
    get name() {
        return commands.TASK_REMOVE;
    }
    get description() {
        return 'Removes the specified Microsoft To Do task';
    }
    constructor() {
        super();
        _TodoTaskRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _TodoTaskRemoveCommand_instances, "m", _TodoTaskRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TodoTaskRemoveCommand_instances, "m", _TodoTaskRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TodoTaskRemoveCommand_instances, "m", _TodoTaskRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeToDoTask(args.options);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the task ${args.options.id} from list ${args.options.listId || args.options.listName}?` });
            if (result) {
                await this.removeToDoTask(args.options);
            }
        }
    }
    async getToDoListId(options) {
        if (options.listName) {
            // Search list by its name
            const requestOptions = {
                url: `${this.resource}/v1.0/me/todo/lists?$filter=displayName eq '${formatting.encodeQueryParameter(options.listName)}'`,
                headers: {
                    accept: "application/json;odata.metadata=none"
                },
                responseType: 'json'
            };
            const response = await request.get(requestOptions);
            return response.value && response.value.length === 1 ? response.value[0].id : undefined;
        }
        return options.listId;
    }
    async removeToDoTask(options) {
        try {
            const toDoListId = await this.getToDoListId(options);
            if (!toDoListId) {
                throw `The list ${options.listName} cannot be found`;
            }
            const requestOptions = {
                url: `${this.resource}/v1.0/me/todo/lists/${toDoListId}/tasks/${options.id}`,
                headers: {
                    accept: "application/json;odata.metadata=none"
                },
                responseType: 'json'
            };
            await request.delete(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_TodoTaskRemoveCommand_instances = new WeakSet(), _TodoTaskRemoveCommand_initTelemetry = function _TodoTaskRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listName: typeof args.options.listName !== 'undefined',
            listId: typeof args.options.listId !== 'undefined',
            force: typeof args.options.force !== 'undefined'
        });
    });
}, _TodoTaskRemoveCommand_initOptions = function _TodoTaskRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '--listName [listName]'
    }, {
        option: '--listId [listId]'
    }, {
        option: '-f, --force'
    });
}, _TodoTaskRemoveCommand_initOptionSets = function _TodoTaskRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['listName', 'listId'] });
};
export default new TodoTaskRemoveCommand();
//# sourceMappingURL=task-remove.js.map