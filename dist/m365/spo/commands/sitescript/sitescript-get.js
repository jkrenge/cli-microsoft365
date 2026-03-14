var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteScriptGetCommand_instances, _SpoSiteScriptGetCommand_initOptions, _SpoSiteScriptGetCommand_initValidators;
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteScriptGetCommand extends SpoCommand {
    get name() {
        return commands.SITESCRIPT_GET;
    }
    get description() {
        return 'Gets information about the specified site script';
    }
    constructor() {
        super();
        _SpoSiteScriptGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteScriptGetCommand_instances, "m", _SpoSiteScriptGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteScriptGetCommand_instances, "m", _SpoSiteScriptGetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoUrl = await spo.getSpoUrl(logger, this.debug);
            const formDigest = await spo.getRequestDigest(spoUrl);
            const requestOptions = {
                url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScriptMetadata`,
                headers: {
                    'X-RequestDigest': formDigest.FormDigestValue,
                    'content-type': 'application/json;charset=utf-8',
                    accept: 'application/json;odata=nometadata'
                },
                data: { id: args.options.id },
                responseType: 'json'
            };
            const response = await request.post(requestOptions);
            response.Content = JSON.parse(response.Content);
            if (args.options.content) {
                await logger.log(response.Content);
                return;
            }
            await logger.log(response);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoSiteScriptGetCommand_instances = new WeakSet(), _SpoSiteScriptGetCommand_initOptions = function _SpoSiteScriptGetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '-c, --content'
    });
}, _SpoSiteScriptGetCommand_initValidators = function _SpoSiteScriptGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
};
export default new SpoSiteScriptGetCommand();
//# sourceMappingURL=sitescript-get.js.map