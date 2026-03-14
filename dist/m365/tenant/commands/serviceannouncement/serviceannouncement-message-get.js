var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TenantServiceAnnouncementMessageGetCommand_instances, _TenantServiceAnnouncementMessageGetCommand_initOptions, _TenantServiceAnnouncementMessageGetCommand_initValidators;
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TenantServiceAnnouncementMessageGetCommand extends GraphCommand {
    get name() {
        return commands.SERVICEANNOUNCEMENT_MESSAGE_GET;
    }
    get description() {
        return 'Retrieves a specified service update message for the tenant';
    }
    constructor() {
        super();
        _TenantServiceAnnouncementMessageGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _TenantServiceAnnouncementMessageGetCommand_instances, "m", _TenantServiceAnnouncementMessageGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TenantServiceAnnouncementMessageGetCommand_instances, "m", _TenantServiceAnnouncementMessageGetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving service update message ${args.options.id}`);
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/admin/serviceAnnouncement/messages/${args.options.id}`,
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
    isValidId(id) {
        return (/MC\d{6}/).test(id);
    }
}
_TenantServiceAnnouncementMessageGetCommand_instances = new WeakSet(), _TenantServiceAnnouncementMessageGetCommand_initOptions = function _TenantServiceAnnouncementMessageGetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    });
}, _TenantServiceAnnouncementMessageGetCommand_initValidators = function _TenantServiceAnnouncementMessageGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!this.isValidId(args.options.id)) {
            return `${args.options.id} is not a valid message ID`;
        }
        return true;
    });
};
export default new TenantServiceAnnouncementMessageGetCommand();
//# sourceMappingURL=serviceannouncement-message-get.js.map