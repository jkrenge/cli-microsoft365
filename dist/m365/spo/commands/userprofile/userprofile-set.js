var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoUserProfileSetCommand_instances, _SpoUserProfileSetCommand_initOptions;
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoUserProfileSetCommand extends SpoCommand {
    get name() {
        return commands.USERPROFILE_SET;
    }
    get description() {
        return 'Sets user profile property for a SharePoint user';
    }
    constructor() {
        super();
        _SpoUserProfileSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoUserProfileSetCommand_instances, "m", _SpoUserProfileSetCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoUrl = await spo.getSpoUrl(logger, this.debug);
            const res = await spo.getRequestDigest(spoUrl);
            const propertyValue = args.options.propertyValue.split(',').map(o => o.trim());
            let propertyType = 'SetSingleValueProfileProperty';
            const data = {
                accountName: `i:0#.f|membership|${args.options.userName}`,
                propertyName: args.options.propertyName
            };
            if (propertyValue.length > 1) {
                propertyType = 'SetMultiValuedProfileProperty';
                data.propertyValues = [...propertyValue];
            }
            else {
                data.propertyValue = propertyValue[0];
            }
            const requestOptions = {
                url: `${spoUrl}/_api/SP.UserProfiles.PeopleManager/${propertyType}`,
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'Content-type': 'application/json;odata=verbose',
                    'X-RequestDigest': res.FormDigestValue
                },
                data: data,
                responseType: 'json'
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoUserProfileSetCommand_instances = new WeakSet(), _SpoUserProfileSetCommand_initOptions = function _SpoUserProfileSetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --userName <userName>'
    }, {
        option: '-n, --propertyName <propertyName>'
    }, {
        option: '-v, --propertyValue <propertyValue>'
    });
};
export default new SpoUserProfileSetCommand();
//# sourceMappingURL=userprofile-set.js.map