var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoUserProfileGetCommand_instances, _SpoUserProfileGetCommand_initOptions, _SpoUserProfileGetCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoUserProfileGetCommand extends SpoCommand {
    get name() {
        return commands.USERPROFILE_GET;
    }
    get description() {
        return 'Gets user profile property for a SharePoint user';
    }
    constructor() {
        super();
        _SpoUserProfileGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoUserProfileGetCommand_instances, "m", _SpoUserProfileGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoUserProfileGetCommand_instances, "m", _SpoUserProfileGetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoUrl = await spo.getSpoUrl(logger, this.debug);
            const userName = `i:0#.f|membership|${args.options.userName}`;
            const requestOptions = {
                url: `${spoUrl}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='${formatting.encodeQueryParameter(`${userName}`)}'`,
                headers: {
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            const res = await request.get(requestOptions);
            if (!args.options.output || cli.shouldTrimOutput(args.options.output)) {
                res.UserProfileProperties = JSON.stringify(res.UserProfileProperties);
            }
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoUserProfileGetCommand_instances = new WeakSet(), _SpoUserProfileGetCommand_initOptions = function _SpoUserProfileGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --userName <userName>'
    });
}, _SpoUserProfileGetCommand_initValidators = function _SpoUserProfileGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid user principal name`;
        }
        return true;
    });
};
export default new SpoUserProfileGetCommand();
//# sourceMappingURL=userprofile-get.js.map