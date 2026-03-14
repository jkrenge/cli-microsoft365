var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraOAuth2GrantRemoveCommand_instances, _EntraOAuth2GrantRemoveCommand_initOptions;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraOAuth2GrantRemoveCommand extends GraphCommand {
    get name() {
        return commands.OAUTH2GRANT_REMOVE;
    }
    get description() {
        return 'Remove specified service principal OAuth2 permissions';
    }
    constructor() {
        super();
        _EntraOAuth2GrantRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraOAuth2GrantRemoveCommand_instances, "m", _EntraOAuth2GrantRemoveCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        const removeOauth2Grant = async () => {
            if (this.verbose) {
                await logger.logToStderr(`Removing OAuth2 permissions...`);
            }
            try {
                const requestOptions = {
                    url: `${this.resource}/v1.0/oauth2PermissionGrants/${formatting.encodeQueryParameter(args.options.grantId)}`,
                    headers: {
                        'accept': 'application/json;odata.metadata=none'
                    },
                    responseType: 'json'
                };
                await request.delete(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeOauth2Grant();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the OAuth2 permissions for ${args.options.grantId}?` });
            if (result) {
                await removeOauth2Grant();
            }
        }
    }
}
_EntraOAuth2GrantRemoveCommand_instances = new WeakSet(), _EntraOAuth2GrantRemoveCommand_initOptions = function _EntraOAuth2GrantRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-i, --grantId <grantId>'
    }, {
        option: '-f, --force'
    });
};
export default new EntraOAuth2GrantRemoveCommand();
//# sourceMappingURL=oauth2grant-remove.js.map