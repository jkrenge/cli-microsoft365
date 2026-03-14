var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteDesignSetCommand_instances, _SpoSiteDesignSetCommand_initTelemetry, _SpoSiteDesignSetCommand_initOptions, _SpoSiteDesignSetCommand_initTypes, _SpoSiteDesignSetCommand_initValidators;
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteDesignSetCommand extends SpoCommand {
    get name() {
        return commands.SITEDESIGN_SET;
    }
    get description() {
        return 'Updates a site design with new values';
    }
    constructor() {
        super();
        _SpoSiteDesignSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteDesignSetCommand_instances, "m", _SpoSiteDesignSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteDesignSetCommand_instances, "m", _SpoSiteDesignSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteDesignSetCommand_instances, "m", _SpoSiteDesignSetCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _SpoSiteDesignSetCommand_instances, "m", _SpoSiteDesignSetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoUrl = await spo.getSpoUrl(logger, this.debug);
            const updateInfo = {
                Id: args.options.id
            };
            if (args.options.title) {
                updateInfo.Title = args.options.title;
            }
            if (args.options.description) {
                updateInfo.Description = args.options.description;
            }
            if (args.options.siteScripts) {
                updateInfo.SiteScriptIds = args.options.siteScripts.split(',').map(i => i.trim());
            }
            if (args.options.previewImageUrl) {
                updateInfo.PreviewImageUrl = args.options.previewImageUrl;
            }
            if (args.options.previewImageAltText) {
                updateInfo.PreviewImageAltText = args.options.previewImageAltText;
            }
            if (args.options.thumbnailUrl) {
                updateInfo.ThumbnailUrl = args.options.thumbnailUrl;
            }
            if (args.options.webTemplate) {
                updateInfo.WebTemplate = args.options.webTemplate === 'TeamSite' ? '64' : '68';
            }
            if (args.options.version) {
                updateInfo.Version = args.options.version;
            }
            if (typeof args.options.isDefault !== 'undefined') {
                updateInfo.IsDefault = args.options.isDefault;
            }
            const requestOptions = {
                url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign`,
                headers: {
                    'content-type': 'application/json;charset=utf-8',
                    accept: 'application/json;odata=nometadata',
                    responseType: 'json'
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
_SpoSiteDesignSetCommand_instances = new WeakSet(), _SpoSiteDesignSetCommand_initTelemetry = function _SpoSiteDesignSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            title: typeof args.options.title !== 'undefined',
            webTemplate: args.options.webTemplate,
            siteScripts: typeof args.options.siteScripts !== 'undefined',
            description: typeof args.options.description !== 'undefined',
            previewImageUrl: typeof args.options.previewImageUrl !== 'undefined',
            previewImageAltText: typeof args.options.previewImageAltText !== 'undefined',
            thumbnailUrl: typeof args.options.thumbnailUrl !== 'undefined',
            version: typeof args.options.version !== 'undefined',
            isDefault: typeof args.options.isDefault
        });
    });
}, _SpoSiteDesignSetCommand_initOptions = function _SpoSiteDesignSetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '-t, --title [title]'
    }, {
        option: '-w, --webTemplate [webTemplate]',
        autocomplete: ['TeamSite', 'CommunicationSite']
    }, {
        option: '-s, --siteScripts [siteScripts]'
    }, {
        option: '-d, --description [description]'
    }, {
        option: '-m, --previewImageUrl [previewImageUrl]'
    }, {
        option: '-a, --previewImageAltText [previewImageAltText]'
    }, {
        option: '--thumbnailUrl [thumbnailUrl]'
    }, {
        option: '-v, --version [version]'
    }, {
        option: '--isDefault [isDefault]',
        autocomplete: ['true', 'false']
    });
}, _SpoSiteDesignSetCommand_initTypes = function _SpoSiteDesignSetCommand_initTypes() {
    this.types.boolean.push('isDefault');
}, _SpoSiteDesignSetCommand_initValidators = function _SpoSiteDesignSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        if (args.options.webTemplate &&
            args.options.webTemplate !== 'TeamSite' &&
            args.options.webTemplate !== 'CommunicationSite') {
            return `${args.options.webTemplate} is not a valid web template type. Allowed values TeamSite|CommunicationSite`;
        }
        if (args.options.siteScripts) {
            const siteScripts = args.options.siteScripts.split(',');
            for (let i = 0; i < siteScripts.length; i++) {
                const trimmedId = siteScripts[i].trim();
                if (!validation.isValidGuid(trimmedId)) {
                    return `${trimmedId} is not a valid GUID`;
                }
            }
        }
        if (args.options.version &&
            typeof args.options.version !== 'number') {
            return `${args.options.version} is not a number`;
        }
        return true;
    });
};
export default new SpoSiteDesignSetCommand();
//# sourceMappingURL=sitedesign-set.js.map