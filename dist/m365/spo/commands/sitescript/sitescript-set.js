var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteScriptSetCommand_instances, _SpoSiteScriptSetCommand_initTelemetry, _SpoSiteScriptSetCommand_initOptions, _SpoSiteScriptSetCommand_initValidators;
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteScriptSetCommand extends SpoCommand {
    get name() {
        return commands.SITESCRIPT_SET;
    }
    get description() {
        return 'Updates existing site script';
    }
    constructor() {
        super();
        _SpoSiteScriptSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteScriptSetCommand_instances, "m", _SpoSiteScriptSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteScriptSetCommand_instances, "m", _SpoSiteScriptSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteScriptSetCommand_instances, "m", _SpoSiteScriptSetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoUrl = await spo.getSpoUrl(logger, this.debug);
            const formDigest = await spo.getRequestDigest(spoUrl);
            const updateInfo = {
                Id: args.options.id
            };
            if (args.options.title) {
                updateInfo.Title = args.options.title;
            }
            if (args.options.description) {
                updateInfo.Description = args.options.description;
            }
            if (args.options.version) {
                updateInfo.Version = parseInt(args.options.version);
            }
            if (args.options.content) {
                updateInfo.Content = args.options.content;
            }
            const requestOptions = {
                url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteScript`,
                headers: {
                    'X-RequestDigest': formDigest.FormDigestValue,
                    'content-type': 'application/json;charset=utf-8',
                    accept: 'application/json;odata=nometadata'
                },
                data: { updateInfo: updateInfo },
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
_SpoSiteScriptSetCommand_instances = new WeakSet(), _SpoSiteScriptSetCommand_initTelemetry = function _SpoSiteScriptSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            title: (!(!args.options.title)).toString(),
            description: (!(!args.options.description)).toString(),
            version: (!(!args.options.version)).toString(),
            content: (!(!args.options.content)).toString()
        });
    });
}, _SpoSiteScriptSetCommand_initOptions = function _SpoSiteScriptSetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '-t, --title [title]'
    }, {
        option: '-d, --description [description]'
    }, {
        option: '-v, --version [version]'
    }, {
        option: '-c, --content [content]'
    });
}, _SpoSiteScriptSetCommand_initValidators = function _SpoSiteScriptSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        if (args.options.version) {
            const version = parseInt(args.options.version);
            if (isNaN(version)) {
                return `${args.options.version} is not a number`;
            }
        }
        if (args.options.content) {
            try {
                JSON.parse(args.options.content);
            }
            catch (e) {
                return `Specified content value is not a valid JSON string. Error: ${e}`;
            }
        }
        return true;
    });
};
export default new SpoSiteScriptSetCommand();
//# sourceMappingURL=sitescript-set.js.map