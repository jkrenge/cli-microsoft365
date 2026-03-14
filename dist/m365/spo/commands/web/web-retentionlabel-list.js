var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoWebRetentionLabelListCommand_instances, _SpoWebRetentionLabelListCommand_initOptions, _SpoWebRetentionLabelListCommand_initValidators;
import { formatting } from '../../../../utils/formatting.js';
import { odata } from "../../../../utils/odata.js";
import { validation } from "../../../../utils/validation.js";
import SpoCommand from "../../../base/SpoCommand.js";
import commands from "../../commands.js";
class SpoWebRetentionLabelListCommand extends SpoCommand {
    get name() {
        return commands.WEB_RETENTIONLABEL_LIST;
    }
    get description() {
        return `Get a list of retention labels that are available on a site.`;
    }
    defaultProperties() {
        return ['TagId', 'TagName'];
    }
    constructor() {
        super();
        _SpoWebRetentionLabelListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoWebRetentionLabelListCommand_instances, "m", _SpoWebRetentionLabelListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoWebRetentionLabelListCommand_instances, "m", _SpoWebRetentionLabelListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving all retention labels that are available on ${args.options.webUrl}...`);
        }
        const requestUrl = `${args.options.webUrl}/_api/SP.CompliancePolicy.SPPolicyStoreProxy.GetAvailableTagsForSite(siteUrl=@a1)?@a1='${formatting.encodeQueryParameter(args.options.webUrl)}'`;
        try {
            const response = await odata.getAllItems(requestUrl);
            await logger.log(response);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoWebRetentionLabelListCommand_instances = new WeakSet(), _SpoWebRetentionLabelListCommand_initOptions = function _SpoWebRetentionLabelListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    });
}, _SpoWebRetentionLabelListCommand_initValidators = function _SpoWebRetentionLabelListCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.webUrl));
};
export default new SpoWebRetentionLabelListCommand();
//# sourceMappingURL=web-retentionlabel-list.js.map