var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TenantServiceAnnouncementHealthListCommand_instances, _TenantServiceAnnouncementHealthListCommand_initTelemetry, _TenantServiceAnnouncementHealthListCommand_initOptions;
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TenantServiceAnnouncementHealthListCommand extends GraphCommand {
    get name() {
        return commands.SERVICEANNOUNCEMENT_HEALTH_LIST;
    }
    get description() {
        return 'This operation provides the health report of all subscribed services for a tenant';
    }
    defaultProperties() {
        return ['id', 'status', 'service'];
    }
    constructor() {
        super();
        _TenantServiceAnnouncementHealthListCommand_instances.add(this);
        __classPrivateFieldGet(this, _TenantServiceAnnouncementHealthListCommand_instances, "m", _TenantServiceAnnouncementHealthListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TenantServiceAnnouncementHealthListCommand_instances, "m", _TenantServiceAnnouncementHealthListCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        try {
            const response = await this.listServiceHealth(args.options);
            await logger.log(response);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async listServiceHealth(options) {
        const requestOptions = {
            url: `${this.resource}/v1.0/admin/serviceAnnouncement/healthOverviews${options.issues && (!options.output || options.output.toLocaleLowerCase() === 'json') ? '?$expand=issues' : ''}`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        const serviceHealthList = response.value;
        if (!serviceHealthList) {
            throw `Error fetching service health`;
        }
        return serviceHealthList;
    }
}
_TenantServiceAnnouncementHealthListCommand_instances = new WeakSet(), _TenantServiceAnnouncementHealthListCommand_initTelemetry = function _TenantServiceAnnouncementHealthListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            issues: typeof args.options.issues !== 'undefined'
        });
    });
}, _TenantServiceAnnouncementHealthListCommand_initOptions = function _TenantServiceAnnouncementHealthListCommand_initOptions() {
    this.options.unshift({ option: '-i, --issues' });
};
export default new TenantServiceAnnouncementHealthListCommand();
//# sourceMappingURL=serviceannouncement-health-list.js.map