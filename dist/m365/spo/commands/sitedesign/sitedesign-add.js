var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteDesignAddCommand_instances, _SpoSiteDesignAddCommand_initTelemetry, _SpoSiteDesignAddCommand_initOptions, _SpoSiteDesignAddCommand_initValidators;
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteDesignAddCommand extends SpoCommand {
    get name() {
        return commands.SITEDESIGN_ADD;
    }
    get description() {
        return 'Adds site design for creating modern sites';
    }
    constructor() {
        super();
        _SpoSiteDesignAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteDesignAddCommand_instances, "m", _SpoSiteDesignAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteDesignAddCommand_instances, "m", _SpoSiteDesignAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteDesignAddCommand_instances, "m", _SpoSiteDesignAddCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoUrl = await spo.getSpoUrl(logger, this.debug);
            const requestDigest = await spo.getRequestDigest(spoUrl);
            const info = {
                Title: args.options.title,
                WebTemplate: args.options.webTemplate === 'TeamSite' ? '64' : '68',
                SiteScriptIds: args.options.siteScripts.split(',').map(i => i.trim())
            };
            if (args.options.description) {
                info.Description = args.options.description;
            }
            if (args.options.previewImageUrl) {
                info.PreviewImageUrl = args.options.previewImageUrl;
            }
            if (args.options.previewImageAltText) {
                info.PreviewImageAltText = args.options.previewImageAltText;
            }
            if (args.options.thumbnailUrl) {
                info.ThumbnailUrl = args.options.thumbnailUrl;
            }
            if (args.options.isDefault) {
                info.IsDefault = true;
            }
            const requestOptions = {
                url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`,
                headers: {
                    'X-RequestDigest': requestDigest.FormDigestValue,
                    'content-type': 'application/json;charset=utf-8',
                    accept: 'application/json;odata=nometadata'
                },
                data: { info: info },
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
_SpoSiteDesignAddCommand_instances = new WeakSet(), _SpoSiteDesignAddCommand_initTelemetry = function _SpoSiteDesignAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            webTemplate: args.options.webTemplate,
            numSiteScripts: args.options.siteScripts.split(',').length,
            description: (!(!args.options.description)).toString(),
            previewImageUrl: (!(!args.options.previewImageUrl)).toString(),
            previewImageAltText: (!(!args.options.previewImageAltText)).toString(),
            thumbnailUrl: (!(!args.options.thumbnailUrl)).toString(),
            isDefault: args.options.isDefault || false
        });
    });
}, _SpoSiteDesignAddCommand_initOptions = function _SpoSiteDesignAddCommand_initOptions() {
    this.options.unshift({
        option: '-t, --title <title>'
    }, {
        option: '-w, --webTemplate <webTemplate>',
        autocomplete: ['TeamSite', 'CommunicationSite']
    }, {
        option: '-s, --siteScripts <siteScripts>'
    }, {
        option: '-d, --description [description]'
    }, {
        option: '-m, --previewImageUrl [previewImageUrl]'
    }, {
        option: '-a, --previewImageAltText [previewImageAltText]'
    }, {
        option: '--thumbnailUrl [thumbnailUrl]'
    }, {
        option: '--isDefault'
    });
}, _SpoSiteDesignAddCommand_initValidators = function _SpoSiteDesignAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.webTemplate !== 'TeamSite' &&
            args.options.webTemplate !== 'CommunicationSite') {
            return `${args.options.webTemplate} is not a valid web template type. Allowed values TeamSite|CommunicationSite`;
        }
        const siteScripts = args.options.siteScripts.split(',');
        for (let i = 0; i < siteScripts.length; i++) {
            const trimmedId = siteScripts[i].trim();
            if (!validation.isValidGuid(trimmedId)) {
                return `${trimmedId} is not a valid GUID`;
            }
        }
        return true;
    });
};
export default new SpoSiteDesignAddCommand();
//# sourceMappingURL=sitedesign-add.js.map