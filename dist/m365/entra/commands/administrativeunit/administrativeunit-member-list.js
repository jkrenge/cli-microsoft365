var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraAdministrativeUnitMemberListCommand_instances, _EntraAdministrativeUnitMemberListCommand_initTelemetry, _EntraAdministrativeUnitMemberListCommand_initOptions, _EntraAdministrativeUnitMemberListCommand_initValidators, _EntraAdministrativeUnitMemberListCommand_initOptionSets;
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { entraAdministrativeUnit } from '../../../../utils/entraAdministrativeUnit.js';
import { validation } from '../../../../utils/validation.js';
class EntraAdministrativeUnitMemberListCommand extends GraphCommand {
    get name() {
        return commands.ADMINISTRATIVEUNIT_MEMBER_LIST;
    }
    get description() {
        return 'Retrieves members (users, groups, or devices) of an administrative unit.';
    }
    defaultProperties() {
        return ['id', 'displayName'];
    }
    constructor() {
        super();
        _EntraAdministrativeUnitMemberListCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraAdministrativeUnitMemberListCommand_instances, "m", _EntraAdministrativeUnitMemberListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraAdministrativeUnitMemberListCommand_instances, "m", _EntraAdministrativeUnitMemberListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraAdministrativeUnitMemberListCommand_instances, "m", _EntraAdministrativeUnitMemberListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraAdministrativeUnitMemberListCommand_instances, "m", _EntraAdministrativeUnitMemberListCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        let administrativeUnitId = args.options.administrativeUnitId;
        try {
            if (args.options.administrativeUnitName) {
                administrativeUnitId = (await entraAdministrativeUnit.getAdministrativeUnitByDisplayName(args.options.administrativeUnitName)).id;
            }
            let results;
            const endpoint = this.getRequestUrl(administrativeUnitId, args.options);
            if (args.options.type) {
                if (args.options.filter) {
                    // While using the filter, we need to specify the ConsistencyLevel header.
                    // Can be refactored when the header is no longer necessary.
                    const requestOptions = {
                        url: endpoint,
                        headers: {
                            accept: 'application/json;odata.metadata=none',
                            ConsistencyLevel: 'eventual'
                        },
                        responseType: 'json'
                    };
                    results = await odata.getAllItems(requestOptions);
                }
                else {
                    results = await odata.getAllItems(endpoint);
                }
            }
            else {
                results = await odata.getAllItems(endpoint, 'minimal');
                results.forEach(c => {
                    const odataType = c['@odata.type'];
                    if (odataType) {
                        c.type = odataType.replace('#microsoft.graph.', '');
                    }
                    delete c['@odata.type'];
                });
            }
            await logger.log(results);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getRequestUrl(administrativeUnitId, options) {
        const queryParameters = [];
        if (options.properties) {
            const allProperties = options.properties.split(',');
            const selectProperties = allProperties.filter(prop => !prop.includes('/'));
            const expandProperties = allProperties.filter(prop => prop.includes('/'));
            if (selectProperties.length > 0) {
                queryParameters.push(`$select=${selectProperties}`);
            }
            if (expandProperties.length > 0) {
                const fieldExpands = expandProperties.map(p => `${p.split('/')[0]}($select=${p.split('/')[1]})`);
                queryParameters.push(`$expand=${fieldExpands.join(',')}`);
            }
        }
        if (options.filter) {
            queryParameters.push(`$filter=${options.filter}`);
            queryParameters.push('$count=true');
        }
        const queryString = queryParameters.length > 0
            ? `?${queryParameters.join('&')}`
            : '';
        return options.type
            ? `${this.resource}/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/microsoft.graph.${options.type}${queryString}`
            : `${this.resource}/v1.0/directory/administrativeUnits/${administrativeUnitId}/members${queryString}`;
    }
}
_EntraAdministrativeUnitMemberListCommand_instances = new WeakSet(), _EntraAdministrativeUnitMemberListCommand_initTelemetry = function _EntraAdministrativeUnitMemberListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            type: typeof args.options.type !== 'undefined',
            properties: typeof args.options.properties !== 'undefined',
            filter: typeof args.options.filter !== 'undefined'
        });
    });
}, _EntraAdministrativeUnitMemberListCommand_initOptions = function _EntraAdministrativeUnitMemberListCommand_initOptions() {
    this.options.unshift({
        option: '-i, --administrativeUnitId [administrativeUnitId]'
    }, {
        option: '-n, --administrativeUnitName [administrativeUnitName]'
    }, {
        option: '-t, --type [type]',
        autocomplete: ['user', 'group', 'device']
    }, {
        option: '-p, --properties [properties]'
    }, {
        option: '-f, --filter [filter]'
    });
}, _EntraAdministrativeUnitMemberListCommand_initValidators = function _EntraAdministrativeUnitMemberListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.administrativeUnitId && !validation.isValidGuid(args.options.administrativeUnitId)) {
            return `${args.options.administrativeUnitId} is not a valid GUID`;
        }
        if (args.options.type) {
            if (['user', 'group', 'device'].every(type => type !== args.options.type)) {
                return `${args.options.type} is not a valid type value. Allowed values user|group|device`;
            }
        }
        if (args.options.filter && !args.options.type) {
            return 'Filter can be specified only if type is set';
        }
        return true;
    });
}, _EntraAdministrativeUnitMemberListCommand_initOptionSets = function _EntraAdministrativeUnitMemberListCommand_initOptionSets() {
    this.optionSets.push({ options: ['administrativeUnitId', 'administrativeUnitName'] });
};
export default new EntraAdministrativeUnitMemberListCommand();
//# sourceMappingURL=administrativeunit-member-list.js.map