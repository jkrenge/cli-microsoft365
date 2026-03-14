var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _GraphSchemaExtensionListCommand_instances, _GraphSchemaExtensionListCommand_initTelemetry, _GraphSchemaExtensionListCommand_initOptions, _GraphSchemaExtensionListCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class GraphSchemaExtensionListCommand extends GraphCommand {
    get name() {
        return commands.SCHEMAEXTENSION_LIST;
    }
    get description() {
        return 'Get a list of schemaExtension objects created in the current tenant, that can be InDevelopment, Available, or Deprecated.';
    }
    constructor() {
        super();
        _GraphSchemaExtensionListCommand_instances.add(this);
        __classPrivateFieldGet(this, _GraphSchemaExtensionListCommand_instances, "m", _GraphSchemaExtensionListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _GraphSchemaExtensionListCommand_instances, "m", _GraphSchemaExtensionListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _GraphSchemaExtensionListCommand_instances, "m", _GraphSchemaExtensionListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const filter = this.getFilter(args.options);
        let url = `${this.resource}/v1.0/schemaExtensions?$select=*${(filter.length > 0 ? '&' + filter : '')}`;
        if (args.options.pageNumber && Number(args.options.pageNumber) > 0) {
            const rowLimit = `&$top=${Number(args.options.pageSize ? args.options.pageSize : 10) * Number(args.options.pageNumber + 1)}`;
            url += rowLimit;
        }
        const requestOptions = {
            url: url,
            headers: {
                accept: 'application/json;odata.metadata=none',
                'content-type': 'application/json'
            },
            responseType: 'json'
        };
        try {
            const res = await request.get(requestOptions);
            if (res.value && res.value.length > 0) {
                const size = args.options.pageSize ? parseInt(args.options.pageSize) : parseInt(res.value.length);
                const result = res.value.slice(-size);
                if (args.options.output !== 'json' && result.length > 1) {
                    await logger.log(result.map((x) => ({
                        id: x.id,
                        description: x.description,
                        targetTypes: x.targetTypes,
                        status: x.status,
                        owner: x.owner,
                        properties: JSON.stringify(x.properties)
                    })));
                }
                else {
                    await logger.log(result);
                }
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getFilter(options) {
        const filters = {};
        const filterOptions = [
            'status',
            'owner'
        ];
        Object.keys(options).forEach(key => {
            if (filterOptions.indexOf(key) !== -1) {
                filters[key] = options[key].replace(/'/g, `''`);
            }
        });
        let filter = Object.keys(filters).map(key => `${key} eq '${filters[key]}'`).join(' and ');
        if (filter.length > 0) {
            filter = '$filter=' + filter;
        }
        return filter;
    }
}
_GraphSchemaExtensionListCommand_instances = new WeakSet(), _GraphSchemaExtensionListCommand_initTelemetry = function _GraphSchemaExtensionListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            status: typeof args.options.status !== 'undefined',
            owner: typeof args.options.owner !== 'undefined',
            pageNumber: typeof args.options.pageNumber !== 'undefined',
            pageSize: typeof args.options.pageSize !== 'undefined'
        });
    });
}, _GraphSchemaExtensionListCommand_initOptions = function _GraphSchemaExtensionListCommand_initOptions() {
    this.options.unshift({
        option: '-s, --status [status]',
        autocomplete: ['Available', 'InDevelopment', 'Deprecated']
    }, {
        option: '--owner [owner]'
    }, {
        option: '-p, --pageSize [pageSize]'
    }, {
        option: '-n, --pageNumber [pageNumber]'
    });
}, _GraphSchemaExtensionListCommand_initValidators = function _GraphSchemaExtensionListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.owner && !validation.isValidGuid(args.options.owner)) {
            return `${args.options.owner} is not a valid GUID`;
        }
        if (args.options.pageNumber && parseInt(args.options.pageNumber) < 1) {
            return 'pageNumber must be a positive number';
        }
        if (args.options.pageSize && parseInt(args.options.pageSize) < 1) {
            return 'pageSize must be a positive number';
        }
        if (args.options.status &&
            ['Available', 'InDevelopment', 'Deprecated'].indexOf(args.options.status) === -1) {
            return `${args.options.status} is not a valid status value. Allowed values are Available|InDevelopment|Deprecated`;
        }
        return true;
    });
};
export default new GraphSchemaExtensionListCommand();
//# sourceMappingURL=schemaextension-list.js.map