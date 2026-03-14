var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _GraphSchemaExtensionSetCommand_instances, _GraphSchemaExtensionSetCommand_initTelemetry, _GraphSchemaExtensionSetCommand_initOptions, _GraphSchemaExtensionSetCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class GraphSchemaExtensionSetCommand extends GraphCommand {
    get name() {
        return commands.SCHEMAEXTENSION_SET;
    }
    get description() {
        return 'Updates a Microsoft Graph schema extension';
    }
    constructor() {
        super();
        _GraphSchemaExtensionSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _GraphSchemaExtensionSetCommand_instances, "m", _GraphSchemaExtensionSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _GraphSchemaExtensionSetCommand_instances, "m", _GraphSchemaExtensionSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _GraphSchemaExtensionSetCommand_instances, "m", _GraphSchemaExtensionSetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Updating schema extension with id '${args.options.id}'...`);
        }
        // The default request data always contains owner
        const data = {
            owner: args.options.owner
        };
        // Add the description to request data if any
        if (args.options.description) {
            if (this.debug) {
                await logger.logToStderr(`Will update description to '${args.options.description}'...`);
            }
            data.description = args.options.description;
        }
        // Add the status to request data if any
        if (args.options.status) {
            if (this.debug) {
                await logger.logToStderr(`Will update status to '${args.options.status}'...`);
            }
            data.status = args.options.status;
        }
        // Add the target types to request data if any
        const targetTypes = args.options.targetTypes
            ? args.options.targetTypes.split(',').map(t => t.trim())
            : [];
        if (targetTypes.length > 0) {
            if (this.debug) {
                await logger.logToStderr(`Will update targetTypes to '${args.options.targetTypes}'...`);
            }
            data.targetTypes = targetTypes;
        }
        // Add the properties to request data if any
        const properties = args.options.properties
            ? JSON.parse(args.options.properties)
            : null;
        if (properties) {
            if (this.debug) {
                await logger.logToStderr(`Will update properties to '${args.options.properties}'...`);
            }
            data.properties = properties;
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/schemaExtensions/${args.options.id}`,
            headers: {
                accept: 'application/json;odata.metadata=none',
                'content-type': 'application/json'
            },
            data,
            responseType: 'json'
        };
        try {
            await request.patch(requestOptions);
            if (this.debug) {
                await logger.logToStderr("Schema extension successfully updated.");
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    validateProperties(propertiesString) {
        let properties;
        try {
            properties = JSON.parse(propertiesString);
        }
        catch {
            return 'The specified properties is not a valid JSON string';
        }
        // If the properties object is not an array
        if (properties.length === undefined) {
            return 'The specified properties JSON string is not an array';
        }
        for (let i = 0; i < properties.length; i++) {
            const property = properties[i];
            if (!property.name) {
                return `Property ${JSON.stringify(property)} misses name`;
            }
            if (!this.isValidPropertyType(property.type)) {
                return `${property.type} is not a valid property type. Valid types are: Binary, Boolean, DateTime, Integer and String`;
            }
        }
        return true;
    }
    isValidPropertyType(propertyType) {
        if (!propertyType) {
            return false;
        }
        return ['Binary', 'Boolean', 'DateTime', 'Integer', 'String'].indexOf(propertyType) > -1;
    }
}
_GraphSchemaExtensionSetCommand_instances = new WeakSet(), _GraphSchemaExtensionSetCommand_initTelemetry = function _GraphSchemaExtensionSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            description: typeof args.options.description !== 'undefined',
            properties: typeof args.options.properties !== 'undefined',
            targetTypes: typeof args.options.targetTypes !== 'undefined',
            status: args.options.status
        });
    });
}, _GraphSchemaExtensionSetCommand_initOptions = function _GraphSchemaExtensionSetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '--owner <owner>'
    }, {
        option: '-d, --description [description]'
    }, {
        option: '-s, --status [status]'
    }, {
        option: '-t, --targetTypes [targetTypes]'
    }, {
        option: '-p, --properties [properties]'
    });
}, _GraphSchemaExtensionSetCommand_initValidators = function _GraphSchemaExtensionSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.owner)) {
            return `The specified owner '${args.options.owner}' is not a valid App Id`;
        }
        if (!args.options.status && !args.options.properties && !args.options.targetTypes && !args.options.description) {
            return `No updates were specified. Please specify at least one argument among --status, --targetTypes, --description or --properties`;
        }
        const validStatusValues = ['Available', 'Deprecated'];
        if (args.options.status && validStatusValues.indexOf(args.options.status) < 0) {
            return `Status option is invalid. Valid statuses are: Available or Deprecated`;
        }
        if (args.options.properties) {
            return this.validateProperties(args.options.properties);
        }
        return true;
    });
};
export default new GraphSchemaExtensionSetCommand();
//# sourceMappingURL=schemaextension-set.js.map