var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TenantServiceAnnouncementHealthGetCommand_instances, _TenantServiceAnnouncementHealthGetCommand_initTelemetry, _TenantServiceAnnouncementHealthGetCommand_initOptions;
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TenantServiceAnnouncementHealthGetCommand extends GraphCommand {
    get name() {
        return commands.SERVICEANNOUNCEMENT_HEALTH_GET;
    }
    get description() {
        return 'This operation provides the health information of a specified service for a tenant';
    }
    constructor() {
        super();
        _TenantServiceAnnouncementHealthGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _TenantServiceAnnouncementHealthGetCommand_instances, "m", _TenantServiceAnnouncementHealthGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TenantServiceAnnouncementHealthGetCommand_instances, "m", _TenantServiceAnnouncementHealthGetCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        try {
            const res = await this.getServiceHealth(args.options);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getServiceHealth(options) {
        const requestOptions = {
            url: `${this.resource}/v1.0/admin/serviceAnnouncement/healthOverviews/${options.serviceName}${options.issues && (!options.output || options.output.toLocaleLowerCase() === 'json') ? '?$expand=issues' : ''}`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        return request.get(requestOptions);
    }
}
_TenantServiceAnnouncementHealthGetCommand_instances = new WeakSet(), _TenantServiceAnnouncementHealthGetCommand_initTelemetry = function _TenantServiceAnnouncementHealthGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            issues: typeof args.options.issues !== 'undefined'
        });
    });
}, _TenantServiceAnnouncementHealthGetCommand_initOptions = function _TenantServiceAnnouncementHealthGetCommand_initOptions() {
    this.options.unshift({ option: '-s, --serviceName <serviceName>' }, { option: '-i, --issues' });
};
export default new TenantServiceAnnouncementHealthGetCommand();
//# sourceMappingURL=serviceannouncement-health-get.js.map