var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TenantSecurityAlertsListCommand_instances, _TenantSecurityAlertsListCommand_initTelemetry, _TenantSecurityAlertsListCommand_initOptions;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TenantSecurityAlertsListCommand extends GraphCommand {
    get name() {
        return commands.SECURITY_ALERTS_LIST;
    }
    get description() {
        return 'Gets the security alerts for a tenant';
    }
    defaultProperties() {
        return ['id', 'title', 'severity'];
    }
    constructor() {
        super();
        _TenantSecurityAlertsListCommand_instances.add(this);
        __classPrivateFieldGet(this, _TenantSecurityAlertsListCommand_instances, "m", _TenantSecurityAlertsListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TenantSecurityAlertsListCommand_instances, "m", _TenantSecurityAlertsListCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        try {
            const res = await this.listAlert(args.options);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async listAlert(options) {
        let queryFilter = '';
        if (options.vendor) {
            let vendorName = options.vendor;
            switch (options.vendor.toLowerCase()) {
                case 'azure security center':
                    vendorName = 'ASC';
                    break;
                case 'microsoft cloud app security':
                    vendorName = 'MCAS';
                    break;
                case 'azure active directory identity protection':
                    vendorName = 'IPC';
            }
            queryFilter = `?$filter=vendorInformation/provider eq '${formatting.encodeQueryParameter(vendorName)}'`;
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/security/alerts${queryFilter}`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        const alertList = response.value;
        if (!alertList) {
            throw `Error fetching security alerts`;
        }
        return alertList;
    }
}
_TenantSecurityAlertsListCommand_instances = new WeakSet(), _TenantSecurityAlertsListCommand_initTelemetry = function _TenantSecurityAlertsListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            vendor: typeof args.options.vendor !== 'undefined'
        });
    });
}, _TenantSecurityAlertsListCommand_initOptions = function _TenantSecurityAlertsListCommand_initOptions() {
    this.options.unshift({ option: '--vendor [vendor]' });
};
export default new TenantSecurityAlertsListCommand();
//# sourceMappingURL=security-alerts-list.js.map