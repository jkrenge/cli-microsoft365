var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteDesignRightsGrantCommand_instances, _SpoSiteDesignRightsGrantCommand_initOptions, _SpoSiteDesignRightsGrantCommand_initValidators;
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteDesignRightsGrantCommand extends SpoCommand {
    get name() {
        return commands.SITEDESIGN_RIGHTS_GRANT;
    }
    get description() {
        return 'Grants access to a site design for one or more principals';
    }
    constructor() {
        super();
        _SpoSiteDesignRightsGrantCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteDesignRightsGrantCommand_instances, "m", _SpoSiteDesignRightsGrantCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteDesignRightsGrantCommand_instances, "m", _SpoSiteDesignRightsGrantCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoUrl = await spo.getSpoUrl(logger, this.debug);
            const requestDigest = await spo.getRequestDigest(spoUrl);
            const grantedRights = '1';
            const requestOptions = {
                url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GrantSiteDesignRights`,
                headers: {
                    'X-RequestDigest': requestDigest.FormDigestValue,
                    'content-type': 'application/json;charset=utf-8',
                    accept: 'application/json;odata=nometadata'
                },
                data: {
                    id: args.options.siteDesignId,
                    principalNames: args.options.principals.split(',').map(p => p.trim()),
                    grantedRights: grantedRights
                },
                responseType: 'json'
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoSiteDesignRightsGrantCommand_instances = new WeakSet(), _SpoSiteDesignRightsGrantCommand_initOptions = function _SpoSiteDesignRightsGrantCommand_initOptions() {
    this.options.unshift({
        option: '-i, --siteDesignId <siteDesignId>'
    }, {
        option: '-p, --principals <principals>'
    }, {
        option: '-r, --rights <rights>',
        autocomplete: ['View']
    });
}, _SpoSiteDesignRightsGrantCommand_initValidators = function _SpoSiteDesignRightsGrantCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.siteDesignId)) {
            return `${args.options.siteDesignId} is not a valid GUID`;
        }
        if (args.options.rights !== 'View') {
            return `${args.options.rights} is not a valid rights value. Allowed values View`;
        }
        return true;
    });
};
export default new SpoSiteDesignRightsGrantCommand();
//# sourceMappingURL=sitedesign-rights-grant.js.map