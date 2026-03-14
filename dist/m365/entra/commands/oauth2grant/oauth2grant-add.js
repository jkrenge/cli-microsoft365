var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraOAuth2GrantAddCommand_instances, _EntraOAuth2GrantAddCommand_initOptions, _EntraOAuth2GrantAddCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraOAuth2GrantAddCommand extends GraphCommand {
    get name() {
        return commands.OAUTH2GRANT_ADD;
    }
    get description() {
        return 'Grant the specified service principal OAuth2 permissions to the specified resource';
    }
    constructor() {
        super();
        _EntraOAuth2GrantAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraOAuth2GrantAddCommand_instances, "m", _EntraOAuth2GrantAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraOAuth2GrantAddCommand_instances, "m", _EntraOAuth2GrantAddCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Granting the service principal specified permissions...`);
        }
        try {
            const requestOptions = {
                url: `${this.resource}/v1.0/oauth2PermissionGrants`,
                headers: {
                    'content-type': 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: {
                    "clientId": args.options.clientId,
                    "consentType": "AllPrincipals",
                    "principalId": null,
                    "resourceId": args.options.resourceId,
                    "scope": args.options.scope
                }
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_EntraOAuth2GrantAddCommand_instances = new WeakSet(), _EntraOAuth2GrantAddCommand_initOptions = function _EntraOAuth2GrantAddCommand_initOptions() {
    this.options.unshift({
        option: '-i, --clientId <clientId>'
    }, {
        option: '-r, --resourceId <resourceId>'
    }, {
        option: '-s, --scope <scope>'
    });
}, _EntraOAuth2GrantAddCommand_initValidators = function _EntraOAuth2GrantAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.clientId)) {
            return `${args.options.clientId} is not a valid GUID`;
        }
        if (!validation.isValidGuid(args.options.resourceId)) {
            return `${args.options.resourceId} is not a valid GUID`;
        }
        return true;
    });
};
export default new EntraOAuth2GrantAddCommand();
//# sourceMappingURL=oauth2grant-add.js.map