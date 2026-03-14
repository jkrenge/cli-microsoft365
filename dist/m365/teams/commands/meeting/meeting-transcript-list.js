var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsMeetingTranscriptListCommand_instances, _TeamsMeetingTranscriptListCommand_initTelemetry, _TeamsMeetingTranscriptListCommand_initOptions, _TeamsMeetingTranscriptListCommand_initValidators, _TeamsMeetingTranscriptListCommand_initOptionSets;
import auth from '../../../../Auth.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsMeetingTranscriptListCommand extends GraphCommand {
    get name() {
        return commands.MEETING_TRANSCRIPT_LIST;
    }
    get description() {
        return 'Lists all transcripts for a given meeting';
    }
    defaultProperties() {
        return ['id', 'createdDateTime'];
    }
    constructor() {
        super();
        _TeamsMeetingTranscriptListCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsMeetingTranscriptListCommand_instances, "m", _TeamsMeetingTranscriptListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsMeetingTranscriptListCommand_instances, "m", _TeamsMeetingTranscriptListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsMeetingTranscriptListCommand_instances, "m", _TeamsMeetingTranscriptListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsMeetingTranscriptListCommand_instances, "m", _TeamsMeetingTranscriptListCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken);
            if (this.verbose) {
                await logger.logToStderr(`Retrieving transcript list for the given meeting...`);
            }
            let requestUrl = `${this.resource}/beta/`;
            if (isAppOnlyAccessToken) {
                if (!args.options.userId && !args.options.userName && !args.options.email) {
                    throw `The option 'userId', 'userName' or 'email' is required when retrieving meeting transcripts list using app only permissions`;
                }
                requestUrl += 'users/';
                if (args.options.userId) {
                    requestUrl += args.options.userId;
                }
                else if (args.options.userName) {
                    requestUrl += args.options.userName;
                }
                else if (args.options.email) {
                    if (this.verbose) {
                        await logger.logToStderr(`Getting user ID for user with email '${args.options.email}'.`);
                    }
                    const userId = await entraUser.getUserIdByEmail(args.options.email);
                    requestUrl += userId;
                }
            }
            else {
                if (args.options.userId || args.options.userName || args.options.email) {
                    throw `The options 'userId', 'userName' and 'email' cannot be used while retrieving meeting transcripts using delegated permissions`;
                }
                requestUrl += `me`;
            }
            requestUrl += `/onlineMeetings/${args.options.meetingId}/transcripts`;
            const res = await odata.getAllItems(requestUrl);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_TeamsMeetingTranscriptListCommand_instances = new WeakSet(), _TeamsMeetingTranscriptListCommand_initTelemetry = function _TeamsMeetingTranscriptListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            email: typeof args.options.email !== 'undefined'
        });
    });
}, _TeamsMeetingTranscriptListCommand_initOptions = function _TeamsMeetingTranscriptListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --userId [userId]'
    }, {
        option: '-n, --userName [userName]'
    }, {
        option: '--email [email]'
    }, {
        option: '-m, --meetingId <meetingId>'
    });
}, _TeamsMeetingTranscriptListCommand_initValidators = function _TeamsMeetingTranscriptListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid Guid`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid user principal name (UPN)`;
        }
        if (args.options.email && !validation.isValidUserPrincipalName(args.options.email)) {
            return `${args.options.email} is not a valid email`;
        }
        return true;
    });
}, _TeamsMeetingTranscriptListCommand_initOptionSets = function _TeamsMeetingTranscriptListCommand_initOptionSets() {
    this.optionSets.push({
        options: ['userId', 'userName', 'email'],
        runsWhen: (args) => args.options.userId || args.options.userName || args.options.email
    });
};
export default new TeamsMeetingTranscriptListCommand();
//# sourceMappingURL=meeting-transcript-list.js.map