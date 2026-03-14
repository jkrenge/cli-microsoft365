var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TenantServiceAnnouncementHealthIssueListCommand_instances, _TenantServiceAnnouncementHealthIssueListCommand_initOptions;
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TenantServiceAnnouncementHealthIssueListCommand extends GraphCommand {
    get name() {
        return commands.SERVICEANNOUNCEMENT_HEALTHISSUE_LIST;
    }
    get description() {
        return 'Gets all service health issues for the tenant';
    }
    defaultProperties() {
        return ['id', 'title'];
    }
    constructor() {
        super();
        _TenantServiceAnnouncementHealthIssueListCommand_instances.add(this);
        __classPrivateFieldGet(this, _TenantServiceAnnouncementHealthIssueListCommand_instances, "m", _TenantServiceAnnouncementHealthIssueListCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        let endpoint = `${this.resource}/v1.0/admin/serviceAnnouncement/issues`;
        if (args.options.service) {
            endpoint += `?$filter=service eq '${formatting.encodeQueryParameter(args.options.service)}'`;
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
_TenantServiceAnnouncementHealthIssueListCommand_instances = new WeakSet(), _TenantServiceAnnouncementHealthIssueListCommand_initOptions = function _TenantServiceAnnouncementHealthIssueListCommand_initOptions() {
    this.options.unshift({
        option: '-s, --service [service]'
    });
};
export default new TenantServiceAnnouncementHealthIssueListCommand();
//# sourceMappingURL=serviceannouncement-healthissue-list.js.map