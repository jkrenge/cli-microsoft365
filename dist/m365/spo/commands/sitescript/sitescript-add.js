var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteScriptAddCommand_instances, _SpoSiteScriptAddCommand_initTelemetry, _SpoSiteScriptAddCommand_initOptions, _SpoSiteScriptAddCommand_initValidators;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteScriptAddCommand extends SpoCommand {
    get name() {
        return commands.SITESCRIPT_ADD;
    }
    get description() {
        return 'Adds site script for use with site designs';
    }
    constructor() {
        super();
        _SpoSiteScriptAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteScriptAddCommand_instances, "m", _SpoSiteScriptAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteScriptAddCommand_instances, "m", _SpoSiteScriptAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteScriptAddCommand_instances, "m", _SpoSiteScriptAddCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoUrl = await spo.getSpoUrl(logger, this.debug);
            const requestDigest = await spo.getRequestDigest(spoUrl);
            const requestOptions = {
                url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteScript(Title=@title, Description=@description)?@title='${formatting.encodeQueryParameter(args.options.title)}'&@description='${formatting.encodeQueryParameter(args.options.description || '')}'`,
                headers: {
                    'X-RequestDigest': requestDigest.FormDigestValue,
                    'content-type': 'application/json;charset=utf-8',
                    accept: 'application/json;odata=nometadata'
                },
                data: JSON.parse(args.options.content),
                responseType: 'json'
            };
            const res = await request.post(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoSiteScriptAddCommand_instances = new WeakSet(), _SpoSiteScriptAddCommand_initTelemetry = function _SpoSiteScriptAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            description: typeof args.options.description !== 'undefined'
        });
    });
}, _SpoSiteScriptAddCommand_initOptions = function _SpoSiteScriptAddCommand_initOptions() {
    this.options.unshift({
        option: '-t, --title <title>'
    }, {
        option: '-c, --content <content>'
    }, {
        option: '-d, --description [description]'
    });
}, _SpoSiteScriptAddCommand_initValidators = function _SpoSiteScriptAddCommand_initValidators() {
    this.validators.push(async (args) => {
        try {
            JSON.parse(args.options.content);
        }
        catch (e) {
            return `Specified content value is not a valid JSON string. Error: ${e}`;
        }
        return true;
    });
};
export default new SpoSiteScriptAddCommand();
//# sourceMappingURL=sitescript-add.js.map