var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraUserListCommand_instances, _a, _EntraUserListCommand_initTelemetry, _EntraUserListCommand_initOptions, _EntraUserListCommand_initValidators, _EntraUserListCommand_initTypes;
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { optionsUtils } from '../../../../utils/optionsUtils.js';
class EntraUserListCommand extends GraphCommand {
    get name() {
        return commands.USER_LIST;
    }
    get description() {
        return 'Lists users matching specified criteria';
    }
    allowUnknownOptions() {
        return true;
    }
    defaultProperties() {
        return ['id', 'displayName', 'mail', 'userPrincipalName'];
    }
    constructor() {
        super();
        _EntraUserListCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraUserListCommand_instances, "m", _EntraUserListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraUserListCommand_instances, "m", _EntraUserListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraUserListCommand_instances, "m", _EntraUserListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraUserListCommand_instances, "m", _EntraUserListCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            let url = `${this.resource}/v1.0/users`;
            if (args.options.properties) {
                const selectProperties = args.options.properties;
                const allSelectProperties = selectProperties.split(',');
                const propertiesWithSlash = allSelectProperties.filter(item => item.includes('/'));
                const fieldExpand = propertiesWithSlash
                    .map(p => `${p.split('/')[0]}($select=${p.split('/')[1]})`)
                    .join(',');
                const expandParam = fieldExpand.length > 0 ? `&$expand=${fieldExpand}` : '';
                const selectParam = allSelectProperties.filter(item => !item.includes('/'));
                url += `?$select=${selectParam}${expandParam}`;
            }
            const filter = this.getFilter(args.options);
            if (filter) {
                url += `${args.options.properties ? '&' : '?'}${filter}`;
            }
            const users = await odata.getAllItems(url);
            await logger.log(users);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getFilter(options) {
        const filters = [];
        const unknownOptions = optionsUtils.getUnknownOptions(options, this.options);
        Object.keys(unknownOptions).forEach(key => {
            if (typeof options[key] === 'boolean') {
                throw `Specify value for the ${key} property`;
            }
            filters.push(`startsWith(${key}, '${formatting.encodeQueryParameter(options[key].toString())}')`);
        });
        if (options.type) {
            filters.push(`userType eq '${options.type}'`);
        }
        if (filters.length > 0) {
            return `$filter=${filters.join(' and ')}`;
        }
        return null;
    }
}
_a = EntraUserListCommand, _EntraUserListCommand_instances = new WeakSet(), _EntraUserListCommand_initTelemetry = function _EntraUserListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            type: typeof args.options.type !== 'undefined',
            properties: typeof args.options.properties !== 'undefined'
        });
    });
}, _EntraUserListCommand_initOptions = function _EntraUserListCommand_initOptions() {
    this.options.unshift({
        option: '--type [type]',
        autocomplete: _a.allowedTypes
    }, {
        option: '-p, --properties [properties]'
    });
}, _EntraUserListCommand_initValidators = function _EntraUserListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.type && !_a.allowedTypes.some(t => t.toLowerCase() === args.options.type.toLowerCase())) {
            return `'${args.options.type}' is not a valid value for option 'type'. Allowed values are: ${_a.allowedTypes.join(', ')}.`;
        }
        return true;
    });
}, _EntraUserListCommand_initTypes = function _EntraUserListCommand_initTypes() {
    this.types.string.push('type', 'properties');
};
EntraUserListCommand.allowedTypes = ['Member', 'Guest'];
export default new EntraUserListCommand();
//# sourceMappingURL=user-list.js.map