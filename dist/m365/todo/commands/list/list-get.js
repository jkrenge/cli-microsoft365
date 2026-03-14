var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TodoListGetCommand_instances, _TodoListGetCommand_initTelemetry, _TodoListGetCommand_initOptions, _TodoListGetCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import commands from '../../commands.js';
class TodoListGetCommand extends GraphDelegatedCommand {
    get name() {
        return commands.LIST_GET;
    }
    get description() {
        return 'Gets a specific list of Microsoft To Do task lists';
    }
    constructor() {
        super();
        _TodoListGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _TodoListGetCommand_instances, "m", _TodoListGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TodoListGetCommand_instances, "m", _TodoListGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TodoListGetCommand_instances, "m", _TodoListGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const item = await this.getList(args.options);
            await logger.log(item);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getList(options) {
        const requestOptions = {
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        if (options.id) {
            requestOptions.url = `${this.resource}/v1.0/me/todo/lists/${options.id}`;
            const result = await request.get(requestOptions);
            return result;
        }
        requestOptions.url = `${this.resource}/v1.0/me/todo/lists?$filter=displayName eq '${formatting.encodeQueryParameter(options.name)}'`;
        const result = await request.get(requestOptions);
        if (result.value.length === 0) {
            throw `The specified list '${options.name}' does not exist.`;
        }
        return result.value[0];
    }
}
_TodoListGetCommand_instances = new WeakSet(), _TodoListGetCommand_initTelemetry = function _TodoListGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined'
        });
    });
}, _TodoListGetCommand_initOptions = function _TodoListGetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-n, --name [name]'
    });
}, _TodoListGetCommand_initOptionSets = function _TodoListGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] });
};
export default new TodoListGetCommand();
//# sourceMappingURL=list-get.js.map