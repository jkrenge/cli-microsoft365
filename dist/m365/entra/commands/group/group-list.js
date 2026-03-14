var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraGroupListCommand_instances, _a, _EntraGroupListCommand_initTelemetry, _EntraGroupListCommand_initOptions, _EntraGroupListCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraGroupListCommand extends GraphCommand {
    get name() {
        return commands.GROUP_LIST;
    }
    get description() {
        return 'Lists all groups defined in Entra ID.';
    }
    defaultProperties() {
        return ['id', 'displayName', 'groupType'];
    }
    constructor() {
        super();
        _EntraGroupListCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraGroupListCommand_instances, "m", _EntraGroupListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraGroupListCommand_instances, "m", _EntraGroupListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraGroupListCommand_instances, "m", _EntraGroupListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            let requestUrl = `${this.resource}/v1.0/groups`;
            let useConsistencyLevelHeader = false;
            if (args.options.type) {
                const groupType = _a.groupTypes.find(g => g.toLowerCase() === args.options.type?.toLowerCase());
                switch (groupType) {
                    case 'microsoft365':
                        requestUrl += `?$filter=groupTypes/any(c:c+eq+'Unified')`;
                        break;
                    case 'security':
                        requestUrl += '?$filter=securityEnabled eq true and mailEnabled eq false';
                        break;
                    case 'distribution':
                        useConsistencyLevelHeader = true;
                        requestUrl += `?$filter=securityEnabled eq false and mailEnabled eq true and not(groupTypes/any(t:t eq 'Unified'))&$count=true`;
                        break;
                    case 'mailEnabledSecurity':
                        useConsistencyLevelHeader = true;
                        requestUrl += `?$filter=securityEnabled eq true and mailEnabled eq true and not(groupTypes/any(t:t eq 'Unified'))&$count=true`;
                        break;
                }
            }
            const queryParameters = [];
            if (args.options.properties) {
                const allProperties = args.options.properties.split(',');
                const selectProperties = allProperties.filter(prop => !prop.includes('/'));
                if (selectProperties.length > 0) {
                    queryParameters.push(`$select=${selectProperties}`);
                }
            }
            const queryString = queryParameters.length > 0
                ? `?${queryParameters.join('&')}`
                : '';
            requestUrl += queryString;
            let groups = [];
            if (useConsistencyLevelHeader) {
                // While using not() function in the filter, we need to specify the ConsistencyLevel header.
                const requestOptions = {
                    url: requestUrl,
                    headers: {
                        accept: 'application/json;odata.metadata=none',
                        ConsistencyLevel: 'eventual'
                    },
                    responseType: 'json'
                };
                groups = await odata.getAllItems(requestOptions);
            }
            else {
                groups = await odata.getAllItems(requestUrl);
            }
            if (cli.shouldTrimOutput(args.options.output)) {
                groups.forEach((group) => {
                    if (group.groupTypes && group.groupTypes.length > 0 && group.groupTypes.includes('Unified')) {
                        group.groupType = 'Microsoft 365';
                    }
                    else if (group.mailEnabled && group.securityEnabled) {
                        group.groupType = 'Mail enabled security';
                    }
                    else if (group.securityEnabled) {
                        group.groupType = 'Security';
                    }
                    else if (group.mailEnabled) {
                        group.groupType = 'Distribution';
                    }
                });
            }
            await logger.log(groups);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_a = EntraGroupListCommand, _EntraGroupListCommand_instances = new WeakSet(), _EntraGroupListCommand_initTelemetry = function _EntraGroupListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            type: typeof args.options.type !== 'undefined',
            properties: typeof args.options.properties !== 'undefined'
        });
    });
}, _EntraGroupListCommand_initOptions = function _EntraGroupListCommand_initOptions() {
    this.options.unshift({
        option: '--type [type]',
        autocomplete: _a.groupTypes
    }, {
        option: '-p, --properties [properties]'
    });
}, _EntraGroupListCommand_initValidators = function _EntraGroupListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.type && _a.groupTypes.every(g => g.toLowerCase() !== args.options.type?.toLowerCase())) {
            return `${args.options.type} is not a valid type value. Allowed values microsoft365|security|distribution|mailEnabledSecurity.`;
        }
        return true;
    });
};
EntraGroupListCommand.groupTypes = ['microsoft365', 'security', 'distribution', 'mailEnabledSecurity'];
export default new EntraGroupListCommand();
//# sourceMappingURL=group-list.js.map