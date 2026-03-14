var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteCommSiteEnableCommand_instances, _SpoSiteCommSiteEnableCommand_initTelemetry, _SpoSiteCommSiteEnableCommand_initOptions, _SpoSiteCommSiteEnableCommand_initOptionSets, _SpoSiteCommSiteEnableCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteCommSiteEnableCommand extends SpoCommand {
    get name() {
        return commands.SITE_COMMSITE_ENABLE;
    }
    get description() {
        return 'Enables communication site features on the specified site';
    }
    constructor() {
        super();
        _SpoSiteCommSiteEnableCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteCommSiteEnableCommand_instances, "m", _SpoSiteCommSiteEnableCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteCommSiteEnableCommand_instances, "m", _SpoSiteCommSiteEnableCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteCommSiteEnableCommand_instances, "m", _SpoSiteCommSiteEnableCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoSiteCommSiteEnableCommand_instances, "m", _SpoSiteCommSiteEnableCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const designPackageId = this.getDesignPackageId(args.options);
        if (this.verbose) {
            await logger.logToStderr(`Enabling communication site with design package '${designPackageId}' at '${args.options.url}'...`);
        }
        try {
            const requestOptions = {
                url: `${args.options.url}/_api/sitepages/communicationsite/enable`,
                headers: {
                    'accept': 'application/json;odata=nometadata'
                },
                data: { designPackageId },
                responseType: 'json'
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getDesignPackageId(options) {
        if (options.designPackageId) {
            return options.designPackageId;
        }
        switch (options.designPackage) {
            case 'Blank':
                return 'f6cc5403-0d63-442e-96c0-285923709ffc';
            case 'Showcase':
                return '6142d2a0-63a5-4ba0-aede-d9fefca2c767';
            case 'Topic':
            default:
                return '96c933ac-3698-44c7-9f4a-5fd17d71af9e';
        }
    }
}
_SpoSiteCommSiteEnableCommand_instances = new WeakSet(), _SpoSiteCommSiteEnableCommand_initTelemetry = function _SpoSiteCommSiteEnableCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            designPackageId: typeof args.options.designPackageId !== 'undefined',
            designPackage: typeof args.options.designPackage !== 'undefined'
        });
    });
}, _SpoSiteCommSiteEnableCommand_initOptions = function _SpoSiteCommSiteEnableCommand_initOptions() {
    this.options.unshift({
        option: '-u, --url <url>'
    }, {
        option: '-i, --designPackageId [designPackageId]'
    }, {
        option: '-p, --designPackage [designPackage]',
        autocomplete: ["Topic", "Showcase", "Blank"]
    });
}, _SpoSiteCommSiteEnableCommand_initOptionSets = function _SpoSiteCommSiteEnableCommand_initOptionSets() {
    this.optionSets.push({
        options: ['designPackageId', 'designPackage'],
        runsWhen: (args) => args.options.designPackageId || args.options.designPackage
    });
}, _SpoSiteCommSiteEnableCommand_initValidators = function _SpoSiteCommSiteEnableCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.designPackageId &&
            !validation.isValidGuid(args.options.designPackageId)) {
            return `${args.options.designPackageId} is not a valid GUID.`;
        }
        if (args.options.designPackage) {
            if (['Topic', 'Showcase', 'Blank'].indexOf(args.options.designPackage) === -1) {
                return `${args.options.designPackage} is not a valid designPackage. Allowed values are Topic|Showcase|Blank`;
            }
        }
        return validation.isValidSharePointUrl(args.options.url);
    });
};
export default new SpoSiteCommSiteEnableCommand();
//# sourceMappingURL=site-commsite-enable.js.map