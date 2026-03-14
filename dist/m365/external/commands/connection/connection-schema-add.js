var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _ExternalConnectionSchemaAddCommand_instances, _ExternalConnectionSchemaAddCommand_initOptions, _ExternalConnectionSchemaAddCommand_initValidators;
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class ExternalConnectionSchemaAddCommand extends GraphCommand {
    get name() {
        return commands.CONNECTION_SCHEMA_ADD;
    }
    get description() {
        return 'Allows the administrator to add a schema to a specific external connection';
    }
    alias() {
        return [commands.EXTERNALCONNECTION_SCHEMA_ADD];
    }
    constructor() {
        super();
        _ExternalConnectionSchemaAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _ExternalConnectionSchemaAddCommand_instances, "m", _ExternalConnectionSchemaAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _ExternalConnectionSchemaAddCommand_instances, "m", _ExternalConnectionSchemaAddCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Adding schema to external connection with id ${args.options.externalConnectionId}`);
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/external/connections/${args.options.externalConnectionId}/schema`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json',
            data: args.options.schema,
            fullResponse: true
        };
        try {
            const res = await request.patch(requestOptions);
            const location = res.headers.location;
            await logger.log(location);
            if (!args.options.wait) {
                return;
            }
            let status;
            do {
                if (this.verbose) {
                    await logger.logToStderr(`Waiting 60 seconds...`);
                }
                // waiting 60s before polling as recommended by Microsoft
                await new Promise(resolve => setTimeout(resolve, 60000));
                if (this.debug) {
                    await logger.logToStderr(`Checking schema operation status...`);
                }
                const operation = await request.get({
                    url: location,
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    },
                    responseType: 'json'
                });
                status = operation.status;
                if (this.verbose) {
                    await logger.logToStderr(`Schema operation status: ${status}`);
                }
                if (status === 'failed') {
                    throw `Provisioning schema failed: ${operation.error?.message}`;
                }
            } while (status === 'inprogress');
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_ExternalConnectionSchemaAddCommand_instances = new WeakSet(), _ExternalConnectionSchemaAddCommand_initOptions = function _ExternalConnectionSchemaAddCommand_initOptions() {
    this.options.unshift({
        option: '-i, --externalConnectionId <externalConnectionId>'
    }, {
        option: '-s, --schema <schema>'
    }, {
        option: '--wait'
    });
}, _ExternalConnectionSchemaAddCommand_initValidators = function _ExternalConnectionSchemaAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.externalConnectionId.length < 3 || args.options.externalConnectionId.length > 32) {
            return 'externalConnectionId must be between 3 and 32 characters in length.';
        }
        const alphaNumericRegEx = /[^\w]|_/g;
        if (alphaNumericRegEx.test(args.options.externalConnectionId)) {
            return 'externalConnectionId must only contain alphanumeric characters.';
        }
        if (args.options.externalConnectionId.length > 9 &&
            args.options.externalConnectionId.startsWith('Microsoft')) {
            return 'ID cannot begin with Microsoft';
        }
        const schemaObject = JSON.parse(args.options.schema);
        if (schemaObject.baseType === undefined || schemaObject.baseType !== 'microsoft.graph.externalItem') {
            return `The schema needs a required property 'baseType' with value 'microsoft.graph.externalItem'`;
        }
        if (!schemaObject.properties || schemaObject.properties.length > 128) {
            return `We need at least one property and a maximum of 128 properties in the schema object`;
        }
        return true;
    });
};
export default new ExternalConnectionSchemaAddCommand();
//# sourceMappingURL=connection-schema-add.js.map