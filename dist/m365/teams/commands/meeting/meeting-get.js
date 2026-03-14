var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsMeetingGetCommand_instances, _TeamsMeetingGetCommand_initTelemetry, _TeamsMeetingGetCommand_initOptions, _TeamsMeetingGetCommand_initValidators;
import auth from '../../../../Auth.js';
import request from '../../../../request.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from "../../../base/GraphCommand.js";
import commands from '../../commands.js';
class TeamsMeetingGetCommand extends GraphCommand {
    get name() {
        return commands.MEETING_GET;
    }
    get description() {
        return 'Get specified meeting details';
    }
    constructor() {
        super();
        _TeamsMeetingGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsMeetingGetCommand_instances, "m", _TeamsMeetingGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsMeetingGetCommand_instances, "m", _TeamsMeetingGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsMeetingGetCommand_instances, "m", _TeamsMeetingGetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken);
        if (isAppOnlyAccessToken) {
            if (!args.options.userId && !args.options.userName && !args.options.email) {
                this.handleError(`The option 'userId', 'userName' or 'email' is required when retrieving meetings using app only permissions`);
            }
        }
        else {
            if (!isAppOnlyAccessToken && (args.options.userId || args.options.userName || args.options.email)) {
                this.handleError(`The options 'userId', 'userName' and 'email' cannot be used when retrieving meetings using delegated permissions`);
            }
        }
        if (this.verbose) {
            await logger.logToStderr(`Retrieving meeting for ${isAppOnlyAccessToken ? 'specific user' : 'currently logged in user'}`);
        }
        try {
            let requestUrl = `${this.resource}/v1.0/`;
            if (isAppOnlyAccessToken) {
                requestUrl += 'users/';
                const userId = await this.getUserId(args.options);
                requestUrl += userId;
            }
            else {
                requestUrl += `me`;
            }
            requestUrl += `/onlineMeetings?$filter=JoinWebUrl eq '${formatting.encodeQueryParameter(args.options.joinUrl)}'`;
            const requestOptions = {
                url: requestUrl,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            const res = await request.get(requestOptions);
            if (res.value.length > 0) {
                await logger.log(res.value[0]);
            }
            else {
                throw `The specified meeting was not found`;
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getUserId(options) {
        if (options.userId) {
            return options.userId;
        }
        if (options.userName) {
            return entraUser.getUserIdByUpn(options.userName);
        }
        return entraUser.getUserIdByEmail(options.email);
    }
}
_TeamsMeetingGetCommand_instances = new WeakSet(), _TeamsMeetingGetCommand_initTelemetry = function _TeamsMeetingGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            email: typeof args.options.email !== 'undefined'
        });
    });
}, _TeamsMeetingGetCommand_initOptions = function _TeamsMeetingGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --userId [userId]'
    }, {
        option: '-n, --userName [userName]'
    }, {
        option: '--email [email]'
    }, {
        option: '-j, --joinUrl <joinUrl>'
    });
}, _TeamsMeetingGetCommand_initValidators = function _TeamsMeetingGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid Guid`;
        }
        return true;
    });
};
export default new TeamsMeetingGetCommand();
//# sourceMappingURL=meeting-get.js.map