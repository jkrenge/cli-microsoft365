var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoKnowledgehubSetCommand_instances, _SpoKnowledgehubSetCommand_initOptions, _SpoKnowledgehubSetCommand_initValidators;
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoKnowledgehubSetCommand extends SpoCommand {
    get name() {
        return commands.KNOWLEDGEHUB_SET;
    }
    get description() {
        return 'Sets the Knowledge Hub Site for your tenant';
    }
    constructor() {
        super();
        _SpoKnowledgehubSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoKnowledgehubSetCommand_instances, "m", _SpoKnowledgehubSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoKnowledgehubSetCommand_instances, "m", _SpoKnowledgehubSetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            const reqDigest = await spo.getRequestDigest(spoAdminUrl);
            if (this.verbose) {
                await logger.logToStderr(`Adding ${args.options.siteUrl} as the Knowledge Hub Site`);
            }
            const requestOptions = {
                url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': reqDigest.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions> <ObjectPath Id="35" ObjectPathId="34" /> <Method Name="SetKnowledgeHubSite" Id="36" ObjectPathId="34"> <Parameters> <Parameter Type="String">${formatting.escapeXml(args.options.siteUrl)}</Parameter> </Parameters> </Method> </Actions> <ObjectPaths> <Constructor Id="34" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /> </ObjectPaths></Request>`
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
    }
}
_SpoKnowledgehubSetCommand_instances = new WeakSet(), _SpoKnowledgehubSetCommand_initOptions = function _SpoKnowledgehubSetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --siteUrl <siteUrl>'
    });
}, _SpoKnowledgehubSetCommand_initValidators = function _SpoKnowledgehubSetCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.siteUrl));
};
export default new SpoKnowledgehubSetCommand();
//# sourceMappingURL=knowledgehub-set.js.map