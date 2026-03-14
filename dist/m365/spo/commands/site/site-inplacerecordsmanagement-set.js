var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteInPlaceRecordsManagementSetCommand_instances, _SpoSiteInPlaceRecordsManagementSetCommand_initTelemetry, _SpoSiteInPlaceRecordsManagementSetCommand_initOptions, _SpoSiteInPlaceRecordsManagementSetCommand_initTypes, _SpoSiteInPlaceRecordsManagementSetCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteInPlaceRecordsManagementSetCommand extends SpoCommand {
    get name() {
        return commands.SITE_INPLACERECORDSMANAGEMENT_SET;
    }
    get description() {
        return 'Activates or deactivates in-place records management for a site collection';
    }
    constructor() {
        super();
        _SpoSiteInPlaceRecordsManagementSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteInPlaceRecordsManagementSetCommand_instances, "m", _SpoSiteInPlaceRecordsManagementSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteInPlaceRecordsManagementSetCommand_instances, "m", _SpoSiteInPlaceRecordsManagementSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteInPlaceRecordsManagementSetCommand_instances, "m", _SpoSiteInPlaceRecordsManagementSetCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _SpoSiteInPlaceRecordsManagementSetCommand_instances, "m", _SpoSiteInPlaceRecordsManagementSetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const requestOptions = {
            url: `${args.options.siteUrl}/_api/site/features/${args.options.enabled ? 'add' : 'remove'}`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            data: {
                featureId: 'da2e115b-07e4-49d9-bb2c-35e93bb9fca9',
                force: true
            },
            responseType: 'json'
        };
        if (this.verbose) {
            await logger.logToStderr(`${args.options.enabled ? 'Activating' : 'Deactivating'} in-place records management for site ${args.options.siteUrl}`);
        }
        try {
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoSiteInPlaceRecordsManagementSetCommand_instances = new WeakSet(), _SpoSiteInPlaceRecordsManagementSetCommand_initTelemetry = function _SpoSiteInPlaceRecordsManagementSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            enabled: args.options.enabled
        });
    });
}, _SpoSiteInPlaceRecordsManagementSetCommand_initOptions = function _SpoSiteInPlaceRecordsManagementSetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --siteUrl <siteUrl>'
    }, {
        option: '--enabled <enabled>',
        autocomplete: ['true', 'false']
    });
}, _SpoSiteInPlaceRecordsManagementSetCommand_initTypes = function _SpoSiteInPlaceRecordsManagementSetCommand_initTypes() {
    this.types.boolean.push('enabled');
}, _SpoSiteInPlaceRecordsManagementSetCommand_initValidators = function _SpoSiteInPlaceRecordsManagementSetCommand_initValidators() {
    this.validators.push(async (args) => {
        return validation.isValidSharePointUrl(args.options.siteUrl);
    });
};
export default new SpoSiteInPlaceRecordsManagementSetCommand();
//# sourceMappingURL=site-inplacerecordsmanagement-set.js.map