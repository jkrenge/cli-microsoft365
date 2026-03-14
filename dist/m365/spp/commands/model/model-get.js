var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SppModelGetCommand_instances, _SppModelGetCommand_initTelemetry, _SppModelGetCommand_initOptions, _SppModelGetCommand_initValidators, _SppModelGetCommand_initOptionSets, _SppModelGetCommand_initTypes;
import { odata } from '../../../../utils/odata.js';
import { spp } from '../../../../utils/spp.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SppModelGetCommand extends SpoCommand {
    get name() {
        return commands.MODEL_GET;
    }
    get description() {
        return 'Retrieves information about a document understanding model';
    }
    constructor() {
        super();
        _SppModelGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SppModelGetCommand_instances, "m", _SppModelGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SppModelGetCommand_instances, "m", _SppModelGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SppModelGetCommand_instances, "m", _SppModelGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SppModelGetCommand_instances, "m", _SppModelGetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SppModelGetCommand_instances, "m", _SppModelGetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            const siteUrl = urlUtil.removeTrailingSlashes(args.options.siteUrl);
            await spp.assertSiteIsContentCenter(siteUrl, logger, this.verbose);
            let result = null;
            if (args.options.title) {
                result = await spp.getModelByTitle(siteUrl, args.options.title, logger, this.verbose);
            }
            else {
                result = await spp.getModelById(siteUrl, args.options.id, logger, this.verbose);
            }
            if (args.options.withPublications) {
                if (this.verbose) {
                    await logger.log(`Retrieving publications for model...`);
                }
                result.Publications = await odata.getAllItems(`${siteUrl}/_api/machinelearning/publications/getbymodeluniqueid('${result.UniqueId}')`);
            }
            await logger.log({
                ...result,
                ConfidenceScore: result.ConfidenceScore ? JSON.parse(result.ConfidenceScore) : null,
                Explanations: result.Explanations ? JSON.parse(result.Explanations) : null,
                Schemas: result.Schemas ? JSON.parse(result.Schemas) : null,
                ModelSettings: result.ModelSettings ? JSON.parse(result.ModelSettings) : null
            });
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SppModelGetCommand_instances = new WeakSet(), _SppModelGetCommand_initTelemetry = function _SppModelGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            title: typeof args.options.title !== 'undefined',
            withPublications: !!args.options.withPublications
        });
    });
}, _SppModelGetCommand_initOptions = function _SppModelGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --siteUrl <siteUrl>'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-t, --title [title]'
    }, {
        option: '--withPublications'
    });
}, _SppModelGetCommand_initValidators = function _SppModelGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID for option 'id'.`;
        }
        return validation.isValidSharePointUrl(args.options.siteUrl);
    });
}, _SppModelGetCommand_initOptionSets = function _SppModelGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'title'] });
}, _SppModelGetCommand_initTypes = function _SppModelGetCommand_initTypes() {
    this.types.string.push('siteUrl', 'id', 'title');
    this.types.boolean.push('withPublications');
};
export default new SppModelGetCommand();
//# sourceMappingURL=model-get.js.map