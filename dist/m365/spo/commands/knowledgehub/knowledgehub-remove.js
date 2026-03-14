var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoKnowledgehubRemoveCommand_instances, _SpoKnowledgehubRemoveCommand_initTelemetry, _SpoKnowledgehubRemoveCommand_initOptions;
import { cli } from '../../../../cli/cli.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoKnowledgehubRemoveCommand extends SpoCommand {
    get name() {
        return commands.KNOWLEDGEHUB_REMOVE;
    }
    get description() {
        return 'Removes the Knowledge Hub Site setting for your tenant';
    }
    constructor() {
        super();
        _SpoKnowledgehubRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoKnowledgehubRemoveCommand_instances, "m", _SpoKnowledgehubRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoKnowledgehubRemoveCommand_instances, "m", _SpoKnowledgehubRemoveCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        const removeKnowledgehub = async () => {
            try {
                const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
                const reqDigest = await spo.getRequestDigest(spoAdminUrl);
                if (this.verbose) {
                    await logger.logToStderr(`Removing Knowledge Hub Site settings from your tenant`);
                }
                const requestOptions = {
                    url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                    headers: {
                        'X-RequestDigest': reqDigest.FormDigestValue
                    },
                    data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="29" ObjectPathId="28"/><Method Name="RemoveKnowledgeHubSite" Id="30" ObjectPathId="28"/></Actions><ObjectPaths><Constructor Id="28" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/></ObjectPaths></Request>`
                };
                const res = await request.post(requestOptions);
                const json = JSON.parse(res);
                const response = json[0];
                if (response.ErrorInfo) {
                    throw response.ErrorInfo.ErrorMessage;
                }
                await logger.log(json[json.length - 1]);
            }
            catch (err) {
                this.handleRejectedPromise(err);
            }
        };
        if (args.options.force) {
            if (this.debug) {
                await logger.logToStderr('Confirmation bypassed by entering confirm option. Removing Knowledge Hub Site setting...');
            }
            await removeKnowledgehub();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove Knowledge Hub Site from your tenant?` });
            if (result) {
                await removeKnowledgehub();
            }
        }
    }
}
_SpoKnowledgehubRemoveCommand_instances = new WeakSet(), _SpoKnowledgehubRemoveCommand_initTelemetry = function _SpoKnowledgehubRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: (!(!args.options.force)).toString()
        });
    });
}, _SpoKnowledgehubRemoveCommand_initOptions = function _SpoKnowledgehubRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-f, --force'
    });
};
export default new SpoKnowledgehubRemoveCommand();
//# sourceMappingURL=knowledgehub-remove.js.map