var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoTenantRecycleBinItemRemoveCommand_instances, _SpoTenantRecycleBinItemRemoveCommand_initTelemetry, _SpoTenantRecycleBinItemRemoveCommand_initOptions, _SpoTenantRecycleBinItemRemoveCommand_initValidators;
import { setTimeout } from 'timers/promises';
import { cli } from '../../../../cli/cli.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoTenantRecycleBinItemRemoveCommand extends SpoCommand {
    get name() {
        return commands.TENANT_RECYCLEBINITEM_REMOVE;
    }
    get description() {
        return 'Removes the specified deleted site collection from tenant recycle bin';
    }
    constructor() {
        super();
        _SpoTenantRecycleBinItemRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoTenantRecycleBinItemRemoveCommand_instances, "m", _SpoTenantRecycleBinItemRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoTenantRecycleBinItemRemoveCommand_instances, "m", _SpoTenantRecycleBinItemRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoTenantRecycleBinItemRemoveCommand_instances, "m", _SpoTenantRecycleBinItemRemoveCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeDeletedSite(logger, args);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the deleted site collection ${args.options.siteUrl} from tenant recycle bin?` });
            if (result) {
                await this.removeDeletedSite(logger, args);
            }
        }
    }
    async removeDeletedSite(logger, args) {
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            const res = await spo.ensureFormDigest(spoAdminUrl, logger, this.context, this.debug);
            if (this.verbose) {
                await logger.logToStderr(`Removing deleted site collection ${args.options.siteUrl}...`);
            }
            const requestOptions = {
                url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': res.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="16" ObjectPathId="15" /><Query Id="17" ObjectPathId="15"><Query SelectAllProperties="false"><Properties><Property Name="PollingInterval" ScalarProperty="true" /><Property Name="IsComplete" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="15" ParentId="1" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.siteUrl)}</Parameter></Parameters></Method><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
            };
            const processQuery = await request.post(requestOptions);
            const json = JSON.parse(processQuery);
            const response = json[0];
            if (response.ErrorInfo) {
                throw response.ErrorInfo.ErrorMessage;
            }
            const operation = json[json.length - 1];
            const isComplete = operation.IsComplete;
            if (!args.options.wait || isComplete) {
                return;
            }
            await setTimeout(operation.PollingInterval);
            await spo.waitUntilFinished({
                operationId: JSON.stringify(operation._ObjectIdentity_),
                siteUrl: spoAdminUrl,
                logger,
                currentContext: this.context,
                debug: this.debug,
                verbose: this.verbose
            });
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoTenantRecycleBinItemRemoveCommand_instances = new WeakSet(), _SpoTenantRecycleBinItemRemoveCommand_initTelemetry = function _SpoTenantRecycleBinItemRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            wait: typeof args.options.wait !== 'undefined',
            force: typeof args.options.force !== 'undefined'
        });
    });
}, _SpoTenantRecycleBinItemRemoveCommand_initOptions = function _SpoTenantRecycleBinItemRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --siteUrl <siteUrl>'
    }, {
        option: '--wait'
    }, {
        option: '-f, --force'
    });
}, _SpoTenantRecycleBinItemRemoveCommand_initValidators = function _SpoTenantRecycleBinItemRemoveCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.siteUrl));
};
export default new SpoTenantRecycleBinItemRemoveCommand();
//# sourceMappingURL=tenant-recyclebinitem-remove.js.map