var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoOrgNewsSiteSetCommand_instances, _SpoOrgNewsSiteSetCommand_initOptions, _SpoOrgNewsSiteSetCommand_initValidators;
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoOrgNewsSiteSetCommand extends SpoCommand {
    get name() {
        return commands.ORGNEWSSITE_SET;
    }
    get description() {
        return 'Marks site as an organizational news site';
    }
    constructor() {
        super();
        _SpoOrgNewsSiteSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoOrgNewsSiteSetCommand_instances, "m", _SpoOrgNewsSiteSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoOrgNewsSiteSetCommand_instances, "m", _SpoOrgNewsSiteSetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            const reqDigest = await spo.getRequestDigest(spoAdminUrl);
            const requestOptions = {
                url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': reqDigest.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="61" ObjectPathId="60" /><Method Name="SetOrgNewsSite" Id="62" ObjectPathId="60"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.url)}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="60" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
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
    }
}
_SpoOrgNewsSiteSetCommand_instances = new WeakSet(), _SpoOrgNewsSiteSetCommand_initOptions = function _SpoOrgNewsSiteSetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --url <url>'
    });
}, _SpoOrgNewsSiteSetCommand_initValidators = function _SpoOrgNewsSiteSetCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.url));
};
export default new SpoOrgNewsSiteSetCommand();
//# sourceMappingURL=orgnewssite-set.js.map