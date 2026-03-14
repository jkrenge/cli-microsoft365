var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _GraphSchemaExtensionGetCommand_instances, _GraphSchemaExtensionGetCommand_initOptions;
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class GraphSchemaExtensionGetCommand extends GraphCommand {
    get name() {
        return commands.SCHEMAEXTENSION_GET;
    }
    get description() {
        return 'Gets the properties of the specified schema extension definition';
    }
    constructor() {
        super();
        _GraphSchemaExtensionGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _GraphSchemaExtensionGetCommand_instances, "m", _GraphSchemaExtensionGetCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Gets the properties of the specified schema extension definition with id '${args.options.id}'...`);
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/schemaExtensions/${args.options.id}`,
            headers: {
                accept: 'application/json;odata.metadata=none',
                'content-type': 'application/json'
            },
            responseType: 'json'
        };
        try {
            const res = await request.get(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_GraphSchemaExtensionGetCommand_instances = new WeakSet(), _GraphSchemaExtensionGetCommand_initOptions = function _GraphSchemaExtensionGetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    });
};
export default new GraphSchemaExtensionGetCommand();
//# sourceMappingURL=schemaextension-get.js.map