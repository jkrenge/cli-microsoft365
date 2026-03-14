var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TodoListSetCommand_instances, _TodoListSetCommand_initTelemetry, _TodoListSetCommand_initOptions, _TodoListSetCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import commands from '../../commands.js';
class TodoListSetCommand extends GraphDelegatedCommand {
    get name() {
        return commands.LIST_SET;
    }
    get description() {
        return 'Updates a Microsoft To Do task list';
    }
    constructor() {
        super();
        _TodoListSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _TodoListSetCommand_instances, "m", _TodoListSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TodoListSetCommand_instances, "m", _TodoListSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TodoListSetCommand_instances, "m", _TodoListSetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const endpoint = `${this.resource}/v1.0`;
        const data = {
            displayName: args.options.newName
        };
        try {
            const listId = await this.getListId(args);
            if (!listId) {
                throw `The list ${args.options.name} cannot be found`;
            }
            const requestOptions = {
                url: `${endpoint}/me/todo/lists/${listId}`,
                headers: {
                    accept: 'application/json;odata.metadata=none',
                    'content-type': 'application/json'
                },
                data,
                responseType: 'json'
            };
            await request.patch(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getListId(args) {
        const endpoint = `${this.resource}/v1.0`;
        if (args.options.id) {
            return args.options.id;
        }
        const requestOptions = {
            url: `${endpoint}/me/todo/lists?$filter=displayName eq '${formatting.encodeQueryParameter(args.options.name)}'`,
            headers: {
                accept: "application/json;odata.metadata=none"
            },
            responseType: 'json'
        };
        return request
            .get(requestOptions)
            .then((response) => response.value && response.value.length === 1 ? response.value[0].id : null);
    }
}
_TodoListSetCommand_instances = new WeakSet(), _TodoListSetCommand_initTelemetry = function _TodoListSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined'
        });
    });
}, _TodoListSetCommand_initOptions = function _TodoListSetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-n, --name [name]'
    }, {
        option: '--newName <newName>'
    });
}, _TodoListSetCommand_initOptionSets = function _TodoListSetCommand_initOptionSets() {
    this.optionSets.push({ options: ['name', 'id'] });
};
export default new TodoListSetCommand();
//# sourceMappingURL=list-set.js.map