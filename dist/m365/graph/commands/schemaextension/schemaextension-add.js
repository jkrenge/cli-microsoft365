var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _GraphSchemaExtensionAddCommand_instances, _GraphSchemaExtensionAddCommand_initOptions, _GraphSchemaExtensionAddCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class GraphSchemaExtensionAddCommand extends GraphCommand {
    get name() {
        return commands.SCHEMAEXTENSION_ADD;
    }
    get description() {
        return 'Creates a Microsoft Graph schema extension';
    }
    constructor() {
        super();
        _GraphSchemaExtensionAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _GraphSchemaExtensionAddCommand_instances, "m", _GraphSchemaExtensionAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _GraphSchemaExtensionAddCommand_instances, "m", _GraphSchemaExtensionAddCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Adding schema extension with id '${args.options.id}'...`);
        }
        const targetTypes = args.options.targetTypes.split(',').map(t => t.trim());
        const properties = JSON.parse(args.options.properties);
        const requestOptions = {
            url: `${this.resource}/v1.0/schemaExtensions`,
            headers: {
                accept: 'application/json;odata.metadata=none',
                'content-type': 'application/json'
            },
            data: {
                id: args.options.id,
                description: args.options.description,
                owner: args.options.owner,
                targetTypes,
                properties
            },
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
    validateProperties(propertiesString) {
        let result = false;
        try {
            const properties = JSON.parse(propertiesString);
            // If the properties object is not an array
            if (properties.length === undefined) {
                result = 'The specified JSON string is not an array';
            }
            else {
                for (let i = 0; i < properties.length; i++) {
                    const property = properties[i];
                    if (!property.name) {
                        result = `Property ${JSON.stringify(property)} misses name`;
                    }
                    if (!this.isValidPropertyType(property.type)) {
                        result = `${property.type} is not a valid property type. Valid types are: Binary, Boolean, DateTime, Integer and String`;
                    }
                }
                if (typeof result !== "string") {
                    result = true;
                }
            }
        }
        catch (e) {
            result = e;
        }
        return result;
    }
    isValidPropertyType(propertyType) {
        if (!propertyType) {
            return false;
        }
        return ['Binary', 'Boolean', 'DateTime', 'Integer', 'String'].indexOf(propertyType) > -1;
    }
}
_GraphSchemaExtensionAddCommand_instances = new WeakSet(), _GraphSchemaExtensionAddCommand_initOptions = function _GraphSchemaExtensionAddCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '-d, --description [description]'
    }, {
        option: '--owner <owner>'
    }, {
        option: '-t, --targetTypes <targetTypes>'
    }, {
        option: '-p, --properties <properties>'
    });
}, _GraphSchemaExtensionAddCommand_initValidators = function _GraphSchemaExtensionAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.owner && !validation.isValidGuid(args.options.owner)) {
            return `The specified owner '${args.options.owner}' is not a valid App Id`;
        }
        return this.validateProperties(args.options.properties);
    });
};
export default new GraphSchemaExtensionAddCommand();
//# sourceMappingURL=schemaextension-add.js.map