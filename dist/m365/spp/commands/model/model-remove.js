var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SppModelRemoveCommand_instances, _SppModelRemoveCommand_initTelemetry, _SppModelRemoveCommand_initOptions, _SppModelRemoveCommand_initValidators, _SppModelRemoveCommand_initOptionSets, _SppModelRemoveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spp } from '../../../../utils/spp.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SppModelRemoveCommand extends SpoCommand {
    get name() {
        return commands.MODEL_REMOVE;
    }
    get description() {
        return 'Deletes a document understanding model';
    }
    constructor() {
        super();
        _SppModelRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SppModelRemoveCommand_instances, "m", _SppModelRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SppModelRemoveCommand_instances, "m", _SppModelRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SppModelRemoveCommand_instances, "m", _SppModelRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SppModelRemoveCommand_instances, "m", _SppModelRemoveCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SppModelRemoveCommand_instances, "m", _SppModelRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (!args.options.force) {
                const confirmationResult = await cli.promptForConfirmation({ message: `Are you sure you want to remove model '${args.options.title ?? args.options.id}'?` });
                if (!confirmationResult) {
                    return;
                }
            }
            if (this.verbose) {
                await logger.log(`Removing model from ${args.options.siteUrl}...`);
            }
            const siteUrl = urlUtil.removeTrailingSlashes(args.options.siteUrl);
            await spp.assertSiteIsContentCenter(siteUrl, logger, this.verbose);
            let requestUrl = `${siteUrl}/_api/machinelearning/models/`;
            if (args.options.title) {
                let requestTitle = args.options.title.toLowerCase();
                if (!requestTitle.endsWith('.classifier')) {
                    requestTitle += '.classifier';
                }
                requestUrl += `getbytitle('${formatting.encodeQueryParameter(requestTitle)}')`;
            }
            else {
                requestUrl += `getbyuniqueid('${args.options.id}')`;
            }
            const requestOptions = {
                url: requestUrl,
                headers: {
                    accept: 'application/json;odata=nometadata',
                    'if-match': '*'
                },
                responseType: 'json'
            };
            const result = await request.delete(requestOptions);
            if (result?.['odata.null'] === true) {
                throw "Model not found.";
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SppModelRemoveCommand_instances = new WeakSet(), _SppModelRemoveCommand_initTelemetry = function _SppModelRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            title: typeof args.options.title !== 'undefined',
            force: !!args.options.force
        });
    });
}, _SppModelRemoveCommand_initOptions = function _SppModelRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --siteUrl <siteUrl>'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-t, --title [title]'
    }, {
        option: '-f, --force'
    });
}, _SppModelRemoveCommand_initValidators = function _SppModelRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID for option 'id'.`;
        }
        return validation.isValidSharePointUrl(args.options.siteUrl);
    });
}, _SppModelRemoveCommand_initOptionSets = function _SppModelRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'title'] });
}, _SppModelRemoveCommand_initTypes = function _SppModelRemoveCommand_initTypes() {
    this.types.string.push('siteUrl', 'id', 'title');
    this.types.boolean.push('force');
};
export default new SppModelRemoveCommand();
//# sourceMappingURL=model-remove.js.map