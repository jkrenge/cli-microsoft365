var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraPolicyListCommand_instances, _a, _EntraPolicyListCommand_initTelemetry, _EntraPolicyListCommand_initOptions, _EntraPolicyListCommand_initValidators;
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
const policyEndPoints = {
    activitybasedtimeout: "activityBasedTimeoutPolicies",
    adminconsentrequest: "adminConsentRequestPolicy",
    appManagement: "appManagementPolicies",
    authenticationflows: "authenticationFlowsPolicy",
    authenticationmethods: "authenticationMethodsPolicy",
    authenticationstrength: "authenticationStrengthPolicies",
    authorization: "authorizationPolicy",
    claimsmapping: "claimsMappingPolicies",
    conditionalaccess: "conditionalAccessPolicies",
    crosstenantaccess: "crossTenantAccessPolicy",
    defaultappmanagement: "defaultAppManagementPolicy",
    deviceregistration: "deviceRegistrationPolicy",
    featurerolloutpolicy: "featureRolloutPolicies",
    homerealmdiscovery: "homeRealmDiscoveryPolicies",
    identitysecuritydefaultsenforcement: "identitySecurityDefaultsEnforcementPolicy",
    permissiongrant: "permissionGrantPolicies",
    rolemanagement: "roleManagementPolicies",
    tokenissuance: "tokenIssuancePolicies",
    tokenlifetime: "tokenLifetimePolicies"
};
class EntraPolicyListCommand extends GraphCommand {
    get name() {
        return commands.POLICY_LIST;
    }
    get description() {
        return 'Returns policies from Entra ID';
    }
    constructor() {
        super();
        _EntraPolicyListCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraPolicyListCommand_instances, "m", _EntraPolicyListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraPolicyListCommand_instances, "m", _EntraPolicyListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraPolicyListCommand_instances, "m", _EntraPolicyListCommand_initValidators).call(this);
    }
    defaultProperties() {
        return ['id', 'displayName', 'isOrganizationDefault'];
    }
    async commandAction(logger, args) {
        const policyType = args.options.type ? args.options.type.toLowerCase() : 'all';
        try {
            if (policyType && policyType !== "all") {
                const policies = await this.getPolicies(policyType);
                await logger.log(policies);
            }
            else {
                const policyTypes = Object.keys(policyEndPoints);
                const results = await Promise.all(policyTypes.map(policyType => this.getPolicies(policyType)));
                let allPolicies = [];
                results.forEach((policies) => {
                    allPolicies = allPolicies.concat(policies);
                });
                await logger.log(allPolicies);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getPolicies(policyType) {
        const endpoint = policyEndPoints[policyType];
        let requestUrl = `${this.resource}/v1.0/policies/${endpoint}`;
        if (endpoint === policyEndPoints.rolemanagement) {
            // roleManagementPolicies endpoint requires $filter query parameter
            requestUrl += `?$filter=scopeId eq '/' and scopeType eq 'DirectoryRole'`;
        }
        const requestOptions = {
            url: requestUrl,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        if (endpoint === policyEndPoints.adminconsentrequest ||
            endpoint === policyEndPoints.authenticationflows ||
            endpoint === policyEndPoints.authenticationmethods ||
            endpoint === policyEndPoints.authorization ||
            endpoint === policyEndPoints.crosstenantaccess ||
            endpoint === policyEndPoints.defaultappmanagement ||
            endpoint === policyEndPoints.deviceregistration ||
            endpoint === policyEndPoints.identitysecuritydefaultsenforcement) {
            return response;
        }
        else {
            return response.value;
        }
    }
}
_a = EntraPolicyListCommand, _EntraPolicyListCommand_instances = new WeakSet(), _EntraPolicyListCommand_initTelemetry = function _EntraPolicyListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            policyType: args.options.type || 'all'
        });
    });
}, _EntraPolicyListCommand_initOptions = function _EntraPolicyListCommand_initOptions() {
    this.options.unshift({
        option: '-t, --type [type]',
        autocomplete: _a.supportedPolicyTypes
    });
}, _EntraPolicyListCommand_initValidators = function _EntraPolicyListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.type) {
            const policyType = args.options.type.toLowerCase();
            if (!_a.supportedPolicyTypes.find(p => p.toLowerCase() === policyType)) {
                return `${args.options.type} is not a valid type. Allowed values are ${_a.supportedPolicyTypes.join(', ')}`;
            }
        }
        return true;
    });
};
EntraPolicyListCommand.supportedPolicyTypes = [
    'activityBasedTimeout',
    'adminConsentRequest',
    'appManagement',
    'authenticationFlows',
    'authenticationMethods',
    'authenticationStrength',
    'authorization',
    'claimsMapping',
    'conditionalAccess',
    'crossTenantAccess',
    'defaultAppManagement',
    'deviceRegistration',
    'featureRolloutPolicy',
    'homeRealmDiscovery',
    'identitySecurityDefaultsEnforcement',
    'permissionGrant',
    'roleManagement',
    'tokenIssuance',
    'tokenLifetime'
];
export default new EntraPolicyListCommand();
//# sourceMappingURL=policy-list.js.map