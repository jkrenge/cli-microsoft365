var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteUnarchiveCommand_instances, _SpoSiteUnarchiveCommand_initTelemetry, _SpoSiteUnarchiveCommand_initOptions, _SpoSiteUnarchiveCommand_initValidators, _SpoSiteUnarchiveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteUnarchiveCommand extends SpoCommand {
    get name() {
        return commands.SITE_UNARCHIVE;
    }
    get description() {
        return 'Unarchives a site collection';
    }
    constructor() {
        super();
        _SpoSiteUnarchiveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteUnarchiveCommand_instances, "m", _SpoSiteUnarchiveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteUnarchiveCommand_instances, "m", _SpoSiteUnarchiveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteUnarchiveCommand_instances, "m", _SpoSiteUnarchiveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoSiteUnarchiveCommand_instances, "m", _SpoSiteUnarchiveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.unarchiveSite(logger, args.options.url);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to unarchive site '${args.options.url}'? You won't be able to rearchive this site for 120 days.` });
            if (result) {
                await this.unarchiveSite(logger, args.options.url);
            }
        }
    }
    async unarchiveSite(logger, url) {
        if (this.verbose) {
            await logger.logToStderr(`Unarchiving site '${url}'...`);
        }
        try {
            const adminCenterUrl = await spo.getSpoAdminUrl(logger, this.verbose);
            const requestDigest = await spo.getRequestDigest(adminCenterUrl);
            const requestOptions = {
                url: `${adminCenterUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': requestDigest.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009">
                <Actions>
                  <ObjectPath Id="2" ObjectPathId="1" />
                  <ObjectPath Id="4" ObjectPathId="3" />
                </Actions>
                <ObjectPaths>
                  <Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" />
                  <Method Id="3" ParentId="1" Name="UnarchiveSiteByUrl">
                    <Parameters>
                      <Parameter Type="String">${url}</Parameter>
                    </Parameters>
                  </Method>
                </ObjectPaths>
              </Request>`
            };
            const response = await request.post(requestOptions);
            const json = JSON.parse(response);
            const responseContent = json[0];
            if (responseContent.ErrorInfo) {
                throw responseContent.ErrorInfo.ErrorMessage;
            }
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
}
_SpoSiteUnarchiveCommand_instances = new WeakSet(), _SpoSiteUnarchiveCommand_initTelemetry = function _SpoSiteUnarchiveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: !!args.options.force
        });
    });
}, _SpoSiteUnarchiveCommand_initOptions = function _SpoSiteUnarchiveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --url <url>'
    }, {
        option: '-f, --force'
    });
}, _SpoSiteUnarchiveCommand_initValidators = function _SpoSiteUnarchiveCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.url));
}, _SpoSiteUnarchiveCommand_initTypes = function _SpoSiteUnarchiveCommand_initTypes() {
    this.types.string.push('url');
    this.types.boolean.push('force');
};
export default new SpoSiteUnarchiveCommand();
//# sourceMappingURL=site-unarchive.js.map