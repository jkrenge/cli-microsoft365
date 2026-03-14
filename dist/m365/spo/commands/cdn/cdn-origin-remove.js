var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoCdnOriginRemoveCommand_instances, _SpoCdnOriginRemoveCommand_initTelemetry, _SpoCdnOriginRemoveCommand_initOptions, _SpoCdnOriginRemoveCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoCdnOriginRemoveCommand extends SpoCommand {
    get name() {
        return commands.CDN_ORIGIN_REMOVE;
    }
    get description() {
        return 'Removes CDN origin for the current SharePoint Online tenant';
    }
    constructor() {
        super();
        _SpoCdnOriginRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoCdnOriginRemoveCommand_instances, "m", _SpoCdnOriginRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoCdnOriginRemoveCommand_instances, "m", _SpoCdnOriginRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoCdnOriginRemoveCommand_instances, "m", _SpoCdnOriginRemoveCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const cdnTypeString = args.options.type || 'Public';
        const cdnType = cdnTypeString === 'Private' ? 1 : 0;
        const removeCdnOrigin = async () => {
            try {
                const tenantId = await spo.getTenantId(logger, this.debug);
                const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
                const reqDigest = await spo.getRequestDigest(spoAdminUrl);
                if (this.verbose) {
                    await logger.logToStderr(`Removing origin ${args.options.origin} from the ${(cdnType === 1 ? 'Private' : 'Public')} CDN. Please wait, this might take a moment...`);
                }
                const requestOptions = {
                    url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                    headers: {
                        'X-RequestDigest': reqDigest.FormDigestValue
                    },
                    data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="RemoveTenantCdnOrigin" Id="33" ObjectPathId="29"><Parameters><Parameter Type="Enum">${cdnType}</Parameter><Parameter Type="String">${formatting.escapeXml(args.options.origin)}</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="29" Name="${tenantId}" /></ObjectPaths></Request>`
                };
                const res = await request.post(requestOptions);
                const json = JSON.parse(res);
                const response = json[0];
                if (response.ErrorInfo) {
                    throw response.ErrorInfo.ErrorMessage;
                }
            }
            catch (err) {
                this.handleRejectedPromise(err);
            }
        };
        if (args.options.force) {
            if (this.debug) {
                await logger.logToStderr('Confirmation suppressed through the confirm option. Removing CDN origin...');
            }
            await removeCdnOrigin();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to delete the ${args.options.origin} CDN origin?` });
            if (result) {
                await removeCdnOrigin();
            }
        }
    }
}
_SpoCdnOriginRemoveCommand_instances = new WeakSet(), _SpoCdnOriginRemoveCommand_initTelemetry = function _SpoCdnOriginRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            cdnType: args.options.type || 'Public',
            force: (!(!args.options.force)).toString()
        });
    });
}, _SpoCdnOriginRemoveCommand_initOptions = function _SpoCdnOriginRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-t, --type [type]',
        autocomplete: ['Public', 'Private']
    }, {
        option: '-r, --origin <origin>'
    }, {
        option: '-f, --force'
    });
}, _SpoCdnOriginRemoveCommand_initValidators = function _SpoCdnOriginRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.type) {
            if (args.options.type !== 'Public' &&
                args.options.type !== 'Private') {
                return `${args.options.type} is not a valid CDN type. Allowed values are Public|Private`;
            }
        }
        return true;
    });
};
export default new SpoCdnOriginRemoveCommand();
//# sourceMappingURL=cdn-origin-remove.js.map