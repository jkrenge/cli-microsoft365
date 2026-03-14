var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraMultitenantSetCommand_instances, _EntraMultitenantSetCommand_initTelemetry, _EntraMultitenantSetCommand_initOptions, _EntraMultitenantSetCommand_initValidators, _EntraMultitenantSetCommand_initTypes;
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraMultitenantSetCommand extends GraphCommand {
    get name() {
        return commands.MULTITENANT_SET;
    }
    get description() {
        return 'Updates the properties of a multitenant organization';
    }
    constructor() {
        super();
        _EntraMultitenantSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraMultitenantSetCommand_instances, "m", _EntraMultitenantSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraMultitenantSetCommand_instances, "m", _EntraMultitenantSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraMultitenantSetCommand_instances, "m", _EntraMultitenantSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraMultitenantSetCommand_instances, "m", _EntraMultitenantSetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr('Updating multitenant organization...');
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/tenantRelationships/multiTenantOrganization`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json',
            data: {
                description: args.options.description,
                displayName: args.options.displayName
            }
        };
        try {
            await request.patch(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_EntraMultitenantSetCommand_instances = new WeakSet(), _EntraMultitenantSetCommand_initTelemetry = function _EntraMultitenantSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            displayName: typeof args.options.displayName !== 'undefined',
            description: typeof args.options.description !== 'undefined'
        });
    });
}, _EntraMultitenantSetCommand_initOptions = function _EntraMultitenantSetCommand_initOptions() {
    this.options.unshift({
        option: '-n, --displayName [displayName]'
    }, {
        option: '-d, --description [description]'
    });
}, _EntraMultitenantSetCommand_initValidators = function _EntraMultitenantSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!args.options.displayName && !args.options.description) {
            return 'Specify either displayName or description or both.';
        }
        return true;
    });
}, _EntraMultitenantSetCommand_initTypes = function _EntraMultitenantSetCommand_initTypes() {
    this.types.string.push('displayName', 'description');
};
export default new EntraMultitenantSetCommand();
//# sourceMappingURL=multitenant-set.js.map