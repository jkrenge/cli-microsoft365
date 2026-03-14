var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraUserPasswordValidateCommand_instances, _EntraUserPasswordValidateCommand_initOptions;
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraUserPasswordValidateCommand extends GraphCommand {
    get name() {
        return commands.USER_PASSWORD_VALIDATE;
    }
    get description() {
        return "Check a user's password against the organization's password validation policy";
    }
    constructor() {
        super();
        _EntraUserPasswordValidateCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraUserPasswordValidateCommand_instances, "m", _EntraUserPasswordValidateCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        try {
            const requestOptions = {
                url: `${this.resource}/beta/users/validatePassword`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                data: {
                    password: args.options.password
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
_EntraUserPasswordValidateCommand_instances = new WeakSet(), _EntraUserPasswordValidateCommand_initOptions = function _EntraUserPasswordValidateCommand_initOptions() {
    this.options.unshift({
        option: '-p, --password <password>'
    });
};
export default new EntraUserPasswordValidateCommand();
//# sourceMappingURL=user-password-validate.js.map