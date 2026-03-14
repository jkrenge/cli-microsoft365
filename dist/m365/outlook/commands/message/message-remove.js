var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _OutlookMessageRemoveCommand_instances, _OutlookMessageRemoveCommand_initTelemetry, _OutlookMessageRemoveCommand_initOptions, _OutlookMessageRemoveCommand_initValidators, _OutlookMessageRemoveCommand_initTypes;
import auth from '../../../../Auth.js';
import request from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';
import { validation } from '../../../../utils/validation.js';
import { formatting } from '../../../../utils/formatting.js';
class OutlookMessageRemoveCommand extends GraphCommand {
    get name() {
        return commands.MESSAGE_REMOVE;
    }
    get description() {
        return 'Permanently removes a specific message from a mailbox';
    }
    constructor() {
        super();
        _OutlookMessageRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _OutlookMessageRemoveCommand_instances, "m", _OutlookMessageRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _OutlookMessageRemoveCommand_instances, "m", _OutlookMessageRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _OutlookMessageRemoveCommand_instances, "m", _OutlookMessageRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _OutlookMessageRemoveCommand_instances, "m", _OutlookMessageRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken);
        let principalUrl = '';
        if (isAppOnlyAccessToken) {
            if (!args.options.userId && !args.options.userName) {
                throw `The option 'userId' or 'userName' is required when removing a message using application permissions.`;
            }
            if (args.options.userId && args.options.userName) {
                throw `Both options 'userId' and 'userName' cannot be used together when removing a message using application permissions.`;
            }
        }
        else {
            if (args.options.userId && args.options.userName) {
                throw `Both options 'userId' and 'userName' cannot be used together when removing a message using delegated permissions.`;
            }
        }
        if (args.options.userId || args.options.userName) {
            principalUrl += `users/${args.options.userId || formatting.encodeQueryParameter(args.options.userName)}`;
        }
        else {
            principalUrl += 'me';
        }
        const removeMessage = async () => {
            try {
                if (this.verbose) {
                    await logger.logToStderr(`Removing message with id '${args.options.id}' using ${isAppOnlyAccessToken ? 'application' : 'delegated'} permissions.`);
                }
                const requestOptions = {
                    url: `${this.resource}/v1.0/${principalUrl}/messages/${args.options.id}`,
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    }
                };
                await request.delete(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeMessage();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove message with id '${args.options.id}'?` });
            if (result) {
                await removeMessage();
            }
        }
    }
}
_OutlookMessageRemoveCommand_instances = new WeakSet(), _OutlookMessageRemoveCommand_initTelemetry = function _OutlookMessageRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            force: !!args.options.force
        });
    });
}, _OutlookMessageRemoveCommand_initOptions = function _OutlookMessageRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '--userId [userId]'
    }, {
        option: '--userName [userName]'
    }, {
        option: '-f, --force'
    });
}, _OutlookMessageRemoveCommand_initValidators = function _OutlookMessageRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `Value '${args.options.userId}' is not a valid GUID for option 'userId'.`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `Value '${args.options.userName}' is not a valid user principal name for option 'userName'.`;
        }
        return true;
    });
}, _OutlookMessageRemoveCommand_initTypes = function _OutlookMessageRemoveCommand_initTypes() {
    this.types.string.push('id', 'userId', 'userName');
    this.types.boolean.push('force');
};
export default new OutlookMessageRemoveCommand();
//# sourceMappingURL=message-remove.js.map