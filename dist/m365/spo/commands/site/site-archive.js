var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteArchiveCommand_instances, _SpoSiteArchiveCommand_initTelemetry, _SpoSiteArchiveCommand_initOptions, _SpoSiteArchiveCommand_initValidators, _SpoSiteArchiveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteArchiveCommand extends SpoCommand {
    get name() {
        return commands.SITE_ARCHIVE;
    }
    get description() {
        return 'Archives a site collection';
    }
    constructor() {
        super();
        _SpoSiteArchiveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteArchiveCommand_instances, "m", _SpoSiteArchiveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteArchiveCommand_instances, "m", _SpoSiteArchiveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteArchiveCommand_instances, "m", _SpoSiteArchiveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoSiteArchiveCommand_instances, "m", _SpoSiteArchiveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const archiveSite = async () => {
            try {
                if (this.verbose) {
                    await logger.logToStderr(`Archiving site ${args.options.url}...`);
                }
                const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.verbose);
                const reqDigest = await spo.getRequestDigest(spoAdminUrl);
                const requestOptions = {
                    url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                    headers: {
                        'X-RequestDigest': reqDigest.FormDigestValue
                    },
                    data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009">
                  <Actions>
                    <ObjectPath Id="2" ObjectPathId="1" />
                    <ObjectPath Id="4" ObjectPathId="3" />
                    <Query Id="5" ObjectPathId="3">
                      <Query SelectAllProperties="true">
                        <Properties />
                      </Query>
                    </Query>
                  </Actions>
                  <ObjectPaths>
                    <Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" />
                    <Method Id="3" ParentId="1" Name="ArchiveSiteByUrl">
                      <Parameters>
                        <Parameter Type="String">${args.options.url}</Parameter>
                      </Parameters>
                    </Method>
                  </ObjectPaths>
                  </Request>`
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
            await archiveSite();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to archive site '${args.options.url}'?` });
            if (result) {
                await archiveSite();
            }
        }
    }
}
_SpoSiteArchiveCommand_instances = new WeakSet(), _SpoSiteArchiveCommand_initTelemetry = function _SpoSiteArchiveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: !!args.options.force
        });
    });
}, _SpoSiteArchiveCommand_initOptions = function _SpoSiteArchiveCommand_initOptions() {
    this.options.unshift({ option: '-u, --url <url>' }, { option: '-f, --force' });
}, _SpoSiteArchiveCommand_initValidators = function _SpoSiteArchiveCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.url));
}, _SpoSiteArchiveCommand_initTypes = function _SpoSiteArchiveCommand_initTypes() {
    this.types.string.push('url');
    this.types.boolean.push('force');
};
export default new SpoSiteArchiveCommand();
//# sourceMappingURL=site-archive.js.map