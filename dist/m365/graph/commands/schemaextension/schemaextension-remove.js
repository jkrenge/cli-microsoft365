var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _GraphSchemaExtensionRemoveCommand_instances, _GraphSchemaExtensionRemoveCommand_initTelemetry, _GraphSchemaExtensionRemoveCommand_initOptions;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class GraphSchemaExtensionRemoveCommand extends GraphCommand {
    get name() {
        return commands.SCHEMAEXTENSION_REMOVE;
    }
    get description() {
        return 'Removes specified Microsoft Graph schema extension';
    }
    constructor() {
        super();
        _GraphSchemaExtensionRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _GraphSchemaExtensionRemoveCommand_instances, "m", _GraphSchemaExtensionRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _GraphSchemaExtensionRemoveCommand_instances, "m", _GraphSchemaExtensionRemoveCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        const removeSchemaExtension = async () => {
            if (this.verbose) {
                await logger.logToStderr(`Removes specified Microsoft Graph schema extension with id '${args.options.id}'...`);
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
                await request.delete(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeSchemaExtension();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the schema extension with ID ${args.options.id}?` });
            if (result) {
                await removeSchemaExtension();
            }
        }
    }
}
_GraphSchemaExtensionRemoveCommand_instances = new WeakSet(), _GraphSchemaExtensionRemoveCommand_initTelemetry = function _GraphSchemaExtensionRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: typeof args.options.force !== 'undefined'
        });
    });
}, _GraphSchemaExtensionRemoveCommand_initOptions = function _GraphSchemaExtensionRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '-f, --force'
    });
};
export default new GraphSchemaExtensionRemoveCommand();
//# sourceMappingURL=schemaextension-remove.js.map