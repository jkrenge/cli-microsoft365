var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TodoListRemoveCommand_instances, _TodoListRemoveCommand_initTelemetry, _TodoListRemoveCommand_initOptions, _TodoListRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import commands from '../../commands.js';
class TodoListRemoveCommand extends GraphDelegatedCommand {
    get name() {
        return commands.LIST_REMOVE;
    }
    get description() {
        return 'Removes a Microsoft To Do task list';
    }
    constructor() {
        super();
        _TodoListRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _TodoListRemoveCommand_instances, "m", _TodoListRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TodoListRemoveCommand_instances, "m", _TodoListRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TodoListRemoveCommand_instances, "m", _TodoListRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeList(args);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the task list ${args.options.id || args.options.name}?` });
            if (result) {
                await this.removeList(args);
            }
        }
    }
    async getListId(args) {
        if (args.options.id) {
            return args.options.id;
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/me/todo/lists?$filter=displayName eq '${formatting.encodeQueryParameter(args.options.name)}'`,
            headers: {
                accept: "application/json;odata.metadata=none"
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        return response.value && response.value.length === 1 ? response.value[0].id : undefined;
    }
    async removeList(args) {
        try {
            const listId = await this.getListId(args);
            if (!listId) {
                throw `The list ${args.options.name} cannot be found`;
            }
            const requestOptions = {
                url: `${this.resource}/v1.0/me/todo/lists/${listId}`,
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
_TodoListRemoveCommand_instances = new WeakSet(), _TodoListRemoveCommand_initTelemetry = function _TodoListRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            name: typeof args.options.name !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            force: typeof args.options.force !== 'undefined'
        });
    });
}, _TodoListRemoveCommand_initOptions = function _TodoListRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name [name]'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-f, --force'
    });
}, _TodoListRemoveCommand_initOptionSets = function _TodoListRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['name', 'id'] });
};
export default new TodoListRemoveCommand();
//# sourceMappingURL=list-remove.js.map