var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsMeetingAddCommand_instances, _TeamsMeetingAddCommand_initTelemetry, _TeamsMeetingAddCommand_initOptions, _TeamsMeetingAddCommand_initValidators;
import auth from '../../../../Auth.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from "../../../base/GraphCommand.js";
import commands from '../../commands.js';
import request from '../../../../request.js';
class TeamsMeetingAddCommand extends GraphCommand {
    get name() {
        return commands.MEETING_ADD;
    }
    get description() {
        return 'Creates a new online meeting';
    }
    constructor() {
        super();
        _TeamsMeetingAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsMeetingAddCommand_instances, "m", _TeamsMeetingAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsMeetingAddCommand_instances, "m", _TeamsMeetingAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsMeetingAddCommand_instances, "m", _TeamsMeetingAddCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken);
            if (isAppOnlyAccessToken && !args.options.organizerEmail) {
                throw `The option 'organizerEmail' is required when creating a meeting using app only permissions`;
            }
            if (!isAppOnlyAccessToken && args.options.organizerEmail) {
                throw `The option 'organizerEmail' is not supported when creating a meeting using delegated permissions`;
            }
            const meeting = await this.createMeeting(logger, args.options);
            await logger.log(meeting);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    /**
     * Creates a new online meeting
     * @param logger
     * @param options
     * @returns MS Graph online meeting response
     */
    async createMeeting(logger, options) {
        let requestUrl = `${this.resource}/v1.0/me`;
        if (options.organizerEmail) {
            if (this.verbose) {
                await logger.logToStderr(`Retrieving Organizer Id...`);
            }
            const organizerId = await entraUser.getUserIdByEmail(options.organizerEmail);
            requestUrl = `${this.resource}/v1.0/users/${organizerId}`;
        }
        if (this.verbose) {
            await logger.logToStderr(`Creating the meeting...`);
        }
        const requestData = {};
        if (options.participantUserNames) {
            const attendees = options.participantUserNames.trim().toLowerCase().split(',').map(p => ({
                upn: p.trim()
            }));
            requestData.participants = { attendees };
        }
        if (options.startTime) {
            requestData.startDateTime = options.startTime;
        }
        if (options.endTime) {
            requestData.endDateTime = options.endTime;
            if (!options.startTime) {
                requestData.startDateTime = new Date().toISOString();
            }
        }
        if (options.subject) {
            requestData.subject = options.subject;
        }
        if (options.recordAutomatically !== undefined) {
            requestData.recordAutomatically = true;
        }
        const requestOptions = {
            headers: {
                accept: 'application/json;odata.metadata=none',
                'content-type': 'application/json'
            },
            responseType: 'json',
            url: `${requestUrl}/onlineMeetings`,
            data: requestData
        };
        return request.post(requestOptions);
    }
}
_TeamsMeetingAddCommand_instances = new WeakSet(), _TeamsMeetingAddCommand_initTelemetry = function _TeamsMeetingAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            startTime: typeof args.options.startTime !== 'undefined',
            endTime: typeof args.options.endTime !== 'undefined',
            subject: typeof args.options.subject !== 'undefined',
            participantUserNames: typeof args.options.participantUserNames !== 'undefined',
            organizerEmail: typeof args.options.organizerEmail !== 'undefined',
            recordAutomatically: !!args.options.recordAutomatically
        });
    });
}, _TeamsMeetingAddCommand_initOptions = function _TeamsMeetingAddCommand_initOptions() {
    this.options.unshift({
        option: '-s, --startTime [startTime]'
    }, {
        option: '-e, --endTime [endTime]'
    }, {
        option: '--subject [subject]'
    }, {
        option: '-p, --participantUserNames [participantUserNames]'
    }, {
        option: '--organizerEmail [organizerEmail]'
    }, {
        option: '-r, --recordAutomatically'
    });
}, _TeamsMeetingAddCommand_initValidators = function _TeamsMeetingAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.startTime && !validation.isValidISODateTime(args.options.startTime)) {
            return `'${args.options.startTime}' is not a valid ISO date string for startTime.`;
        }
        if (args.options.endTime && !validation.isValidISODateTime(args.options.endTime)) {
            return `'${args.options.endTime}' is not a valid ISO date string for endTime.`;
        }
        if (args.options.startTime && args.options.endTime && new Date(args.options.startTime) >= new Date(args.options.endTime)) {
            return 'The startTime value must be before endTime.';
        }
        if (args.options.startTime && new Date() >= new Date(args.options.startTime)) {
            return 'The startTime value must be in the future.';
        }
        if (args.options.endTime && new Date() >= new Date(args.options.endTime)) {
            return 'The endTime value must be in the future.';
        }
        if (args.options.participantUserNames) {
            const isValidUPNArrayResult = validation.isValidUserPrincipalNameArray(args.options.participantUserNames);
            if (isValidUPNArrayResult !== true) {
                return `The following user principal names are invalid for the option 'participantUserNames': ${isValidUPNArrayResult}.`;
            }
        }
        if (args.options.organizerEmail && !validation.isValidUserPrincipalName(args.options.organizerEmail)) {
            return `'${args.options.organizerEmail}' is not a valid email for organizerEmail.`;
        }
        return true;
    });
};
export default new TeamsMeetingAddCommand();
//# sourceMappingURL=meeting-add.js.map