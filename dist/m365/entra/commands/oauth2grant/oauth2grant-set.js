var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraOAuth2GrantSetCommand_instances, _EntraOAuth2GrantSetCommand_initOptions;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraOAuth2GrantSetCommand extends GraphCommand {
    get name() {
        return commands.OAUTH2GRANT_SET;
    }
    get description() {
        return 'Update OAuth2 permissions for the service principal';
    }
    constructor() {
        super();
        _EntraOAuth2GrantSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraOAuth2GrantSetCommand_instances, "m", _EntraOAuth2GrantSetCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Updating OAuth2 permissions...`);
        }
        try {
            const requestOptions = {
                url: `${this.resource}/v1.0/oauth2PermissionGrants/${formatting.encodeQueryParameter(args.options.grantId)}`,
                headers: {
                    'content-type': 'application/json'
                },
                responseType: 'json',
                data: {
                    "scope": args.options.scope
                }
            };
            await request.patch(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_EntraOAuth2GrantSetCommand_instances = new WeakSet(), _EntraOAuth2GrantSetCommand_initOptions = function _EntraOAuth2GrantSetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --grantId <grantId>'
    }, {
        option: '-s, --scope <scope>'
    });
};
export default new EntraOAuth2GrantSetCommand();
//# sourceMappingURL=oauth2grant-set.js.map