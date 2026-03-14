var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteDesignRightsRevokeCommand_instances, _SpoSiteDesignRightsRevokeCommand_initTelemetry, _SpoSiteDesignRightsRevokeCommand_initOptions, _SpoSiteDesignRightsRevokeCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteDesignRightsRevokeCommand extends SpoCommand {
    get name() {
        return commands.SITEDESIGN_RIGHTS_REVOKE;
    }
    get description() {
        return 'Revokes access from a site design for one or more principals';
    }
    constructor() {
        super();
        _SpoSiteDesignRightsRevokeCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteDesignRightsRevokeCommand_instances, "m", _SpoSiteDesignRightsRevokeCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteDesignRightsRevokeCommand_instances, "m", _SpoSiteDesignRightsRevokeCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteDesignRightsRevokeCommand_instances, "m", _SpoSiteDesignRightsRevokeCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const revokePermissions = async () => {
            try {
                const spoUrl = await spo.getSpoUrl(logger, this.debug);
                const requestDigest = await spo.getRequestDigest(spoUrl);
                const requestOptions = {
                    url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.RevokeSiteDesignRights`,
                    headers: {
                        'X-RequestDigest': requestDigest.FormDigestValue,
                        'content-type': 'application/json;charset=utf-8',
                        accept: 'application/json;odata=nometadata'
                    },
                    data: {
                        id: args.options.siteDesignId,
                        principalNames: args.options.principals.split(',').map(p => p.trim())
                    },
                    responseType: 'json'
                };
                await request.post(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await revokePermissions();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to revoke access to site design ${args.options.siteDesignId} from the specified users?` });
            if (result) {
                await revokePermissions();
            }
        }
    }
}
_SpoSiteDesignRightsRevokeCommand_instances = new WeakSet(), _SpoSiteDesignRightsRevokeCommand_initTelemetry = function _SpoSiteDesignRightsRevokeCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: args.options.force || false
        });
    });
}, _SpoSiteDesignRightsRevokeCommand_initOptions = function _SpoSiteDesignRightsRevokeCommand_initOptions() {
    this.options.unshift({
        option: '-i, --siteDesignId <siteDesignId>'
    }, {
        option: '-p, --principals <principals>'
    }, {
        option: '-f, --force'
    });
}, _SpoSiteDesignRightsRevokeCommand_initValidators = function _SpoSiteDesignRightsRevokeCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.siteDesignId)) {
            return `${args.options.siteDesignId} is not a valid GUID`;
        }
        return true;
    });
};
export default new SpoSiteDesignRightsRevokeCommand();
//# sourceMappingURL=sitedesign-rights-revoke.js.map