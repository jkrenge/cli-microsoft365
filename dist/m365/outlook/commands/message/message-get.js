var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _OutlookMessageGetCommand_instances, _OutlookMessageGetCommand_initTelemetry, _OutlookMessageGetCommand_initOptions;
import auth from '../../../../Auth.js';
import request from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { mailSenderWhitelist } from '../../../../utils/mailSenderWhitelist.js';
class OutlookMessageGetCommand extends GraphCommand {
    get name() {
        return commands.MESSAGE_GET;
    }
    get description() {
        return 'Retrieves specified message';
    }
    constructor() {
        super();
        _OutlookMessageGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _OutlookMessageGetCommand_instances, "m", _OutlookMessageGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _OutlookMessageGetCommand_instances, "m", _OutlookMessageGetCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        try {
            const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken);
            if (this.verbose) {
                await logger.logToStderr(`Retrieving message with id ${args.options.id} using ${isAppOnlyAccessToken ? 'app only permissions' : 'delegated permissions'}`);
            }
            let requestUrl = '';
            if (isAppOnlyAccessToken) {
                if (!args.options.userId && !args.options.userName) {
                    throw `The option 'userId' or 'userName' is required when retrieving an email using app only credentials`;
                }
                if (args.options.userId && args.options.userName) {
                    throw `Both options 'userId' and 'userName' cannot be set when retrieving an email using app only credentials`;
                }
                requestUrl += `users/${args.options.userId ? args.options.userId : args.options.userName}`;
            }
            else {
                if (args.options.userId && args.options.userName) {
                    throw `Both options 'userId' and 'userName' cannot be set when retrieving an email using delegated credentials`;
                }
                if (args.options.userId || args.options.userName) {
                    requestUrl += `users/${args.options.userId ? args.options.userId : args.options.userName}`;
                }
                else {
                    requestUrl += 'me';
                }
            }
            requestUrl += `/messages/${args.options.id}`;
            const requestOptions = {
                url: `${this.resource}/v1.0/${requestUrl}`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            let res = await request.get(requestOptions);
            if (mailSenderWhitelist.shouldFilterInInteractiveMode()) {
                const filterResult = mailSenderWhitelist.filterResponse(res);
                res = filterResult.filteredResponse;
                await mailSenderWhitelist.logFilteredSummary(logger, filterResult.filteredCount);
            }
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_OutlookMessageGetCommand_instances = new WeakSet(), _OutlookMessageGetCommand_initTelemetry = function _OutlookMessageGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined'
        });
    });
}, _OutlookMessageGetCommand_initOptions = function _OutlookMessageGetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '--userId [userId]'
    }, {
        option: '--userName [userName]'
    });
};
export default new OutlookMessageGetCommand();
//# sourceMappingURL=message-get.js.map