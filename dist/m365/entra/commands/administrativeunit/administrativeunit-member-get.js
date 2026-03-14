var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraAdministrativeUnitMemberGetCommand_instances, _EntraAdministrativeUnitMemberGetCommand_initTelemetry, _EntraAdministrativeUnitMemberGetCommand_initOptions, _EntraAdministrativeUnitMemberGetCommand_initValidators, _EntraAdministrativeUnitMemberGetCommand_initOptionSets;
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { entraAdministrativeUnit } from '../../../../utils/entraAdministrativeUnit.js';
import request from '../../../../request.js';
class EntraAdministrativeUnitMemberGetCommand extends GraphCommand {
    get name() {
        return commands.ADMINISTRATIVEUNIT_MEMBER_GET;
    }
    get description() {
        return 'Retrieve a specific member (user, group, or device) of an administrative unit';
    }
    constructor() {
        super();
        _EntraAdministrativeUnitMemberGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraAdministrativeUnitMemberGetCommand_instances, "m", _EntraAdministrativeUnitMemberGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraAdministrativeUnitMemberGetCommand_instances, "m", _EntraAdministrativeUnitMemberGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraAdministrativeUnitMemberGetCommand_instances, "m", _EntraAdministrativeUnitMemberGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraAdministrativeUnitMemberGetCommand_instances, "m", _EntraAdministrativeUnitMemberGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        let administrativeUnitId = args.options.administrativeUnitId;
        try {
            if (args.options.administrativeUnitName) {
                if (this.verbose) {
                    await logger.logToStderr(`Retrieving Administrative Unit Id...`);
                }
                administrativeUnitId = (await entraAdministrativeUnit.getAdministrativeUnitByDisplayName(args.options.administrativeUnitName)).id;
            }
            const url = this.getRequestUrl(administrativeUnitId, args.options.id, args.options);
            const requestOptions = {
                url: url,
                headers: {
                    accept: 'application/json;odata.metadata=minimal'
                },
                responseType: 'json'
            };
            const result = await request.get(requestOptions);
            const odataType = result['@odata.type'];
            if (odataType) {
                result.type = odataType.replace('#microsoft.graph.', '');
            }
            delete result['@odata.type'];
            delete result['@odata.context'];
            await logger.log(result);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getRequestUrl(administrativeUnitId, memberId, options) {
        const queryParameters = [];
        if (options.properties) {
            const allProperties = options.properties.split(',');
            const selectProperties = allProperties.filter(prop => !prop.includes('/'));
            const expandProperties = allProperties.filter(prop => prop.includes('/'));
            if (selectProperties.length > 0) {
                queryParameters.push(`$select=${selectProperties}`);
            }
            if (expandProperties.length > 0) {
                const fieldExpands = expandProperties.map(p => {
                    const properties = p.split('/');
                    return `${properties[0]}($select=${properties[1]})`;
                });
                queryParameters.push(`$expand=${fieldExpands.join(',')}`);
            }
        }
        const queryString = queryParameters.length > 0
            ? `?${queryParameters.join('&')}`
            : '';
        return `${this.resource}/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/${memberId}${queryString}`;
    }
}
_EntraAdministrativeUnitMemberGetCommand_instances = new WeakSet(), _EntraAdministrativeUnitMemberGetCommand_initTelemetry = function _EntraAdministrativeUnitMemberGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            administrativeUnitId: typeof args.options.administrativeUnitId !== 'undefined',
            administrativeUnitName: typeof args.options.administrativeUnitName !== 'undefined',
            properties: typeof args.options.properties !== 'undefined'
        });
    });
}, _EntraAdministrativeUnitMemberGetCommand_initOptions = function _EntraAdministrativeUnitMemberGetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '-u, --administrativeUnitId [administrativeUnitId]'
    }, {
        option: '-n, --administrativeUnitName [administrativeUnitName]'
    }, {
        option: '-p, --properties [properties]'
    });
}, _EntraAdministrativeUnitMemberGetCommand_initValidators = function _EntraAdministrativeUnitMemberGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        if (args.options.administrativeUnitId && !validation.isValidGuid(args.options.administrativeUnitId)) {
            return `${args.options.administrativeUnitId} is not a valid GUID`;
        }
        return true;
    });
}, _EntraAdministrativeUnitMemberGetCommand_initOptionSets = function _EntraAdministrativeUnitMemberGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['administrativeUnitId', 'administrativeUnitName'] });
};
export default new EntraAdministrativeUnitMemberGetCommand();
//# sourceMappingURL=administrativeunit-member-get.js.map