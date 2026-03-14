var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteDesignRightsListCommand_instances, _SpoSiteDesignRightsListCommand_initOptions, _SpoSiteDesignRightsListCommand_initValidators;
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteDesignRightsListCommand extends SpoCommand {
    get name() {
        return commands.SITEDESIGN_RIGHTS_LIST;
    }
    get description() {
        return 'Gets a list of principals that have access to a site design';
    }
    constructor() {
        super();
        _SpoSiteDesignRightsListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteDesignRightsListCommand_instances, "m", _SpoSiteDesignRightsListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteDesignRightsListCommand_instances, "m", _SpoSiteDesignRightsListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoUrl = await spo.getSpoUrl(logger, this.debug);
            const requestDigest = await spo.getRequestDigest(spoUrl);
            const requestOptions = {
                url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRights`,
                headers: {
                    'X-RequestDigest': requestDigest.FormDigestValue,
                    'content-type': 'application/json;charset=utf-8',
                    accept: 'application/json;odata=nometadata'
                },
                data: { id: args.options.siteDesignId },
                responseType: 'json'
            };
            const res = await request.post(requestOptions);
            await logger.log(res.value.map(p => {
                p.Rights = p.Rights === "1" ? "View" : p.Rights;
                return p;
            }));
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoSiteDesignRightsListCommand_instances = new WeakSet(), _SpoSiteDesignRightsListCommand_initOptions = function _SpoSiteDesignRightsListCommand_initOptions() {
    this.options.unshift({
        option: '-i, --siteDesignId <siteDesignId>'
    });
}, _SpoSiteDesignRightsListCommand_initValidators = function _SpoSiteDesignRightsListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.siteDesignId)) {
            return `${args.options.siteDesignId} is not a valid GUID`;
        }
        return true;
    });
};
export default new SpoSiteDesignRightsListCommand();
//# sourceMappingURL=sitedesign-rights-list.js.map