var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoHubSiteRegisterCommand_instances, _SpoHubSiteRegisterCommand_initOptions, _SpoHubSiteRegisterCommand_initValidators;
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoHubSiteRegisterCommand extends SpoCommand {
    get name() {
        return commands.HUBSITE_REGISTER;
    }
    get description() {
        return 'Registers the specified site collection as a hub site';
    }
    constructor() {
        super();
        _SpoHubSiteRegisterCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoHubSiteRegisterCommand_instances, "m", _SpoHubSiteRegisterCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoHubSiteRegisterCommand_instances, "m", _SpoHubSiteRegisterCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const reqDigest = await spo.getRequestDigest(args.options.siteUrl);
            const requestOptions = {
                url: `${args.options.siteUrl}/_api/site/RegisterHubSite`,
                headers: {
                    'X-RequestDigest': reqDigest.FormDigestValue,
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            const res = await request.post(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoHubSiteRegisterCommand_instances = new WeakSet(), _SpoHubSiteRegisterCommand_initOptions = function _SpoHubSiteRegisterCommand_initOptions() {
    this.options.unshift({
        option: '-u, --siteUrl <siteUrl>'
    });
}, _SpoHubSiteRegisterCommand_initValidators = function _SpoHubSiteRegisterCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.siteUrl));
};
export default new SpoHubSiteRegisterCommand();
//# sourceMappingURL=hubsite-register.js.map