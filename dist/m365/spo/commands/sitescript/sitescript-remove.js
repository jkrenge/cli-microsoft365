var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteScriptRemoveCommand_instances, _SpoSiteScriptRemoveCommand_initTelemetry, _SpoSiteScriptRemoveCommand_initOptions, _SpoSiteScriptRemoveCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteScriptRemoveCommand extends SpoCommand {
    get name() {
        return commands.SITESCRIPT_REMOVE;
    }
    get description() {
        return 'Removes the specified site script';
    }
    constructor() {
        super();
        _SpoSiteScriptRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteScriptRemoveCommand_instances, "m", _SpoSiteScriptRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteScriptRemoveCommand_instances, "m", _SpoSiteScriptRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteScriptRemoveCommand_instances, "m", _SpoSiteScriptRemoveCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeSiteScript(logger, args.options.id);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the site script ${args.options.id}?` });
            if (result) {
                await this.removeSiteScript(logger, args.options.id);
            }
        }
    }
    async removeSiteScript(logger, id) {
        try {
            const spoUrl = await spo.getSpoUrl(logger, this.debug);
            const formDigest = await spo.getRequestDigest(spoUrl);
            const requestOptions = {
                url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.DeleteSiteScript`,
                headers: {
                    'X-RequestDigest': formDigest.FormDigestValue,
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
_SpoSiteScriptRemoveCommand_instances = new WeakSet(), _SpoSiteScriptRemoveCommand_initTelemetry = function _SpoSiteScriptRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: args.options.force || false
        });
    });
}, _SpoSiteScriptRemoveCommand_initOptions = function _SpoSiteScriptRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '-f, --force'
    });
}, _SpoSiteScriptRemoveCommand_initValidators = function _SpoSiteScriptRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
};
export default new SpoSiteScriptRemoveCommand();
//# sourceMappingURL=sitescript-remove.js.map