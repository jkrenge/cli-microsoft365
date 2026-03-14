var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraOAuth2GrantListCommand_instances, _EntraOAuth2GrantListCommand_initOptions, _EntraOAuth2GrantListCommand_initValidators;
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraOAuth2GrantListCommand extends GraphCommand {
    get name() {
        return commands.OAUTH2GRANT_LIST;
    }
    get description() {
        return 'Lists OAuth2 permission grants for the specified service principal';
    }
    defaultProperties() {
        return ['objectId', 'resourceId', 'scope'];
    }
    constructor() {
        super();
        _EntraOAuth2GrantListCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraOAuth2GrantListCommand_instances, "m", _EntraOAuth2GrantListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraOAuth2GrantListCommand_instances, "m", _EntraOAuth2GrantListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving list of OAuth grants for the service principal...`);
        }
        try {
            const res = await odata.getAllItems(`${this.resource}/v1.0/oauth2PermissionGrants?$filter=clientId eq '${formatting.encodeQueryParameter(args.options.spObjectId)}'`);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_EntraOAuth2GrantListCommand_instances = new WeakSet(), _EntraOAuth2GrantListCommand_initOptions = function _EntraOAuth2GrantListCommand_initOptions() {
    this.options.unshift({
        option: '-i, --spObjectId <spObjectId>'
    });
}, _EntraOAuth2GrantListCommand_initValidators = function _EntraOAuth2GrantListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.spObjectId)) {
            return `${args.options.spObjectId} is not a valid GUID`;
        }
        return true;
    });
};
export default new EntraOAuth2GrantListCommand();
//# sourceMappingURL=oauth2grant-list.js.map