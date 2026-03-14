var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TenantServiceAnnouncementMessageListCommand_instances, _TenantServiceAnnouncementMessageListCommand_initTelemetry, _TenantServiceAnnouncementMessageListCommand_initOptions;
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TenantServiceAnnouncementMessageListCommand extends GraphCommand {
    get name() {
        return commands.SERVICEANNOUNCEMENT_MESSAGE_LIST;
    }
    get description() {
        return 'Gets all service update messages for the tenant';
    }
    defaultProperties() {
        return ['id', 'title'];
    }
    constructor() {
        super();
        _TenantServiceAnnouncementMessageListCommand_instances.add(this);
        __classPrivateFieldGet(this, _TenantServiceAnnouncementMessageListCommand_instances, "m", _TenantServiceAnnouncementMessageListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TenantServiceAnnouncementMessageListCommand_instances, "m", _TenantServiceAnnouncementMessageListCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        let endpoint = `${this.resource}/v1.0/admin/serviceAnnouncement/messages`;
        if (args.options.service) {
            endpoint += `?$filter=services/any(c:c+eq+'${formatting.encodeQueryParameter(args.options.service)}')`;
        }
        try {
            const items = await odata.getAllItems(endpoint);
            await logger.log(items);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_TenantServiceAnnouncementMessageListCommand_instances = new WeakSet(), _TenantServiceAnnouncementMessageListCommand_initTelemetry = function _TenantServiceAnnouncementMessageListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            service: typeof args.options.service !== 'undefined'
        });
    });
}, _TenantServiceAnnouncementMessageListCommand_initOptions = function _TenantServiceAnnouncementMessageListCommand_initOptions() {
    this.options.unshift({
        option: '-s, --service [service]'
    });
};
export default new TenantServiceAnnouncementMessageListCommand();
//# sourceMappingURL=serviceannouncement-message-list.js.map