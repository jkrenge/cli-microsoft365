var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TodoListAddCommand_instances, _TodoListAddCommand_initOptions;
import request from '../../../../request.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import commands from '../../commands.js';
class TodoListAddCommand extends GraphDelegatedCommand {
    get name() {
        return commands.LIST_ADD;
    }
    get description() {
        return 'Adds a new Microsoft To Do task list';
    }
    constructor() {
        super();
        _TodoListAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _TodoListAddCommand_instances, "m", _TodoListAddCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        const data = {
            displayName: args.options.name
        };
        const requestOptions = {
            url: `${this.resource}/v1.0/me/todo/lists`,
            headers: {
                accept: 'application/json;odata.metadata=none',
                'content-type': 'application/json'
            },
            data,
            responseType: 'json'
        };
        try {
            const res = await request.post(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_TodoListAddCommand_instances = new WeakSet(), _TodoListAddCommand_initOptions = function _TodoListAddCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name <name>'
    });
};
export default new TodoListAddCommand();
//# sourceMappingURL=list-add.js.map