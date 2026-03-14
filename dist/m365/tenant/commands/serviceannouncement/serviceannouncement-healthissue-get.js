var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TenantServiceAnnouncementHealthIssueGetCommand_instances, _TenantServiceAnnouncementHealthIssueGetCommand_initOptions;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TenantServiceAnnouncementHealthIssueGetCommand extends GraphCommand {
    get name() {
        return commands.SERVICEANNOUNCEMENT_HEALTHISSUE_GET;
    }
    get description() {
        return 'Gets a specified service health issue for tenant';
    }
    constructor() {
        super();
        _TenantServiceAnnouncementHealthIssueGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _TenantServiceAnnouncementHealthIssueGetCommand_instances, "m", _TenantServiceAnnouncementHealthIssueGetCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        const requestOptions = {
            url: `${this.resource}/v1.0/admin/serviceAnnouncement/issues/${formatting.encodeQueryParameter(args.options.id)}`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        try {
            const res = await request.get(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_TenantServiceAnnouncementHealthIssueGetCommand_instances = new WeakSet(), _TenantServiceAnnouncementHealthIssueGetCommand_initOptions = function _TenantServiceAnnouncementHealthIssueGetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    });
};
export default new TenantServiceAnnouncementHealthIssueGetCommand();
//# sourceMappingURL=serviceannouncement-healthissue-get.js.map