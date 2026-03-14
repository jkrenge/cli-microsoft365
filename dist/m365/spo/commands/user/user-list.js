var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoUserListCommand_instances, _SpoUserListCommand_initOptions, _SpoUserListCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoUserListCommand extends SpoCommand {
    get name() {
        return commands.USER_LIST;
    }
    get description() {
        return 'Lists all the users within specific web';
    }
    defaultProperties() {
        return ['Id', 'Title', 'LoginName'];
    }
    constructor() {
        super();
        _SpoUserListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoUserListCommand_instances, "m", _SpoUserListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoUserListCommand_instances, "m", _SpoUserListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving users from web ${args.options.webUrl}...`);
        }
        const requestOptions = {
            url: `${args.options.webUrl}/_api/web/siteusers`,
            method: 'GET',
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        try {
            const users = await request.get(requestOptions);
            await logger.log(users.value);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoUserListCommand_instances = new WeakSet(), _SpoUserListCommand_initOptions = function _SpoUserListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    });
}, _SpoUserListCommand_initValidators = function _SpoUserListCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.webUrl));
};
export default new SpoUserListCommand();
//# sourceMappingURL=user-list.js.map