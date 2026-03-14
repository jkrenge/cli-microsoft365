var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteDesignRemoveCommand_instances, _SpoSiteDesignRemoveCommand_initTelemetry, _SpoSiteDesignRemoveCommand_initOptions, _SpoSiteDesignRemoveCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteDesignRemoveCommand extends SpoCommand {
    get name() {
        return commands.SITEDESIGN_REMOVE;
    }
    get description() {
        return 'Removes the specified site design';
    }
    constructor() {
        super();
        _SpoSiteDesignRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteDesignRemoveCommand_instances, "m", _SpoSiteDesignRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteDesignRemoveCommand_instances, "m", _SpoSiteDesignRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteDesignRemoveCommand_instances, "m", _SpoSiteDesignRemoveCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeSiteDesign(logger, args.options.id);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the site design ${args.options.id}?` });
            if (result) {
                await this.removeSiteDesign(logger, args.options.id);
            }
        }
    }
    async removeSiteDesign(logger, id) {
        try {
            const spoUrl = await spo.getSpoUrl(logger, this.debug);
            const requestDigest = await spo.getRequestDigest(spoUrl);
            const requestOptions = {
                url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.DeleteSiteDesign`,
                headers: {
                    'X-RequestDigest': requestDigest.FormDigestValue,
                    'content-type': 'application/json;charset=utf-8',
                    accept: 'application/json;odata=nometadata'
                },
                data: { id: id },
                responseType: 'json'
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoSiteDesignRemoveCommand_instances = new WeakSet(), _SpoSiteDesignRemoveCommand_initTelemetry = function _SpoSiteDesignRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: args.options.force || false
        });
    });
}, _SpoSiteDesignRemoveCommand_initOptions = function _SpoSiteDesignRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '-f, --force'
    });
}, _SpoSiteDesignRemoveCommand_initValidators = function _SpoSiteDesignRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
};
export default new SpoSiteDesignRemoveCommand();
//# sourceMappingURL=sitedesign-remove.js.map