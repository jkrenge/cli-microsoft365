var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsMeetingAttendancereportGetCommand_instances, _TeamsMeetingAttendancereportGetCommand_initTelemetry, _TeamsMeetingAttendancereportGetCommand_initOptions, _TeamsMeetingAttendancereportGetCommand_initTypes, _TeamsMeetingAttendancereportGetCommand_initValidators, _TeamsMeetingAttendancereportGetCommand_initOptionSets;
import auth from '../../../../Auth.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from "../../../base/GraphCommand.js";
import commands from '../../commands.js';
import { entraUser } from '../../../../utils/entraUser.js';
import request from '../../../../request.js';
class TeamsMeetingAttendancereportGetCommand extends GraphCommand {
    get name() {
        return commands.MEETING_ATTENDANCEREPORT_GET;
    }
    get description() {
        return 'Gets attendance report for a given meeting';
    }
    constructor() {
        super();
        _TeamsMeetingAttendancereportGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsMeetingAttendancereportGetCommand_instances, "m", _TeamsMeetingAttendancereportGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsMeetingAttendancereportGetCommand_instances, "m", _TeamsMeetingAttendancereportGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsMeetingAttendancereportGetCommand_instances, "m", _TeamsMeetingAttendancereportGetCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _TeamsMeetingAttendancereportGetCommand_instances, "m", _TeamsMeetingAttendancereportGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsMeetingAttendancereportGetCommand_instances, "m", _TeamsMeetingAttendancereportGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken);
            if (isAppOnlyAccessToken && !args.options.userId && !args.options.userName && !args.options.email) {
                throw `The option 'userId', 'userName' or 'email' is required when retrieving meeting attendance report using app only permissions.`;
            }
            else if (!isAppOnlyAccessToken && (args.options.userId || args.options.userName || args.options.email)) {
                throw `The options 'userId', 'userName' and 'email' cannot be used when retrieving meeting attendance report using delegated permissions.`;
            }
            if (this.verbose) {
                await logger.logToStderr(`Retrieving attendance report for ${isAppOnlyAccessToken ? `specific user ${args.options.userId || args.options.userName || args.options.email}.` : 'currently logged in user'}.`);
            }
            let userUrl = '';
            if (isAppOnlyAccessToken) {
                const userId = await this.getUserId(args.options);
                userUrl += `users/${userId}`;
            }
            else {
                userUrl += 'me';
            }
            const requestOptions = {
                url: `${this.resource}/v1.0/${userUrl}/onlineMeetings/${args.options.meetingId}/attendanceReports/${args.options.id}?$expand=attendanceRecords`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            const attendanceReport = await request.get(requestOptions);
            await logger.log(attendanceReport);
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
_TeamsMeetingAttendancereportGetCommand_instances = new WeakSet(), _TeamsMeetingAttendancereportGetCommand_initTelemetry = function _TeamsMeetingAttendancereportGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            email: typeof args.options.email !== 'undefined'
        });
    });
}, _TeamsMeetingAttendancereportGetCommand_initOptions = function _TeamsMeetingAttendancereportGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --userId [userId]'
    }, {
        option: '-n, --userName [userName]'
    }, {
        option: '--email [email]'
    }, {
        option: '-m, --meetingId <meetingId>'
    }, {
        option: '-i, --id <id>'
    });
}, _TeamsMeetingAttendancereportGetCommand_initTypes = function _TeamsMeetingAttendancereportGetCommand_initTypes() {
    this.types.string.push('userId', 'userName', 'email', 'meetingId', 'id');
}, _TeamsMeetingAttendancereportGetCommand_initValidators = function _TeamsMeetingAttendancereportGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID for option 'id'.`;
        }
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID for option 'userId'.`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid UPN.`;
        }
        if (args.options.email && !validation.isValidUserPrincipalName(args.options.email)) {
            return `${args.options.email} is not a valid email.`;
        }
        return true;
    });
}, _TeamsMeetingAttendancereportGetCommand_initOptionSets = function _TeamsMeetingAttendancereportGetCommand_initOptionSets() {
    this.optionSets.push({
        options: ['userId', 'userName', 'email'],
        runsWhen: (args) => args.options.userId || args.options.userName || args.options.email
    });
};
export default new TeamsMeetingAttendancereportGetCommand();
//# sourceMappingURL=meeting-attendancereport-get.js.map