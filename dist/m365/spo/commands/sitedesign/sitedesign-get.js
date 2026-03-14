var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteDesignGetCommand_instances, _SpoSiteDesignGetCommand_initTelemetry, _SpoSiteDesignGetCommand_initOptions, _SpoSiteDesignGetCommand_initValidators, _SpoSiteDesignGetCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteDesignGetCommand extends SpoCommand {
    get name() {
        return commands.SITEDESIGN_GET;
    }
    get description() {
        return 'Gets information about the specified site design';
    }
    constructor() {
        super();
        _SpoSiteDesignGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteDesignGetCommand_instances, "m", _SpoSiteDesignGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteDesignGetCommand_instances, "m", _SpoSiteDesignGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteDesignGetCommand_instances, "m", _SpoSiteDesignGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoSiteDesignGetCommand_instances, "m", _SpoSiteDesignGetCommand_initOptionSets).call(this);
    }
    async getSiteDesignId(args, spoUrl) {
        if (args.options.id) {
            return args.options.id;
        }
        const requestOptions = {
            url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesigns`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const response = await request.post(requestOptions);
        const matchingSiteDesigns = response.value.filter(x => x.Title === args.options.title);
        if (matchingSiteDesigns.length === 0) {
            throw `The specified site design does not exist`;
        }
        if (matchingSiteDesigns.length > 1) {
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', matchingSiteDesigns);
            const result = await cli.handleMultipleResultsFound(`Multiple site designs with title '${args.options.title}' found.`, resultAsKeyValuePair);
            return result.Id;
        }
        return matchingSiteDesigns[0].Id;
    }
    async commandAction(logger, args) {
        try {
            const spoUrl = await spo.getSpoUrl(logger, this.debug);
            const siteDesignId = await this.getSiteDesignId(args, spoUrl);
            const requestOptions = {
                url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignMetadata`,
                headers: {
                    'content-type': 'application/json;charset=utf-8',
                    accept: 'application/json;odata=nometadata'
                },
                data: { id: siteDesignId },
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
_SpoSiteDesignGetCommand_instances = new WeakSet(), _SpoSiteDesignGetCommand_initTelemetry = function _SpoSiteDesignGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            title: typeof args.options.title !== 'undefined'
        });
    });
}, _SpoSiteDesignGetCommand_initOptions = function _SpoSiteDesignGetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '--title [title]'
    });
}, _SpoSiteDesignGetCommand_initValidators = function _SpoSiteDesignGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
}, _SpoSiteDesignGetCommand_initOptionSets = function _SpoSiteDesignGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'title'] });
};
export default new SpoSiteDesignGetCommand();
//# sourceMappingURL=sitedesign-get.js.map