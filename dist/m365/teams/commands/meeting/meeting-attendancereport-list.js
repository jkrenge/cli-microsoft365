var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsMeetingAttendancereportListCommand_instances, _TeamsMeetingAttendancereportListCommand_initTelemetry, _TeamsMeetingAttendancereportListCommand_initOptions, _TeamsMeetingAttendancereportListCommand_initValidators;
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import entraUserGetCommand from '../../../entra/commands/user/user-get.js';
import GraphCommand from "../../../base/GraphCommand.js";
import commands from '../../commands.js';
class TeamsMeetingAttendancereportListCommand extends GraphCommand {
    get name() {
        return commands.MEETING_ATTENDANCEREPORT_LIST;
    }
    get description() {
        return 'Lists all attendance reports for a given meeting';
    }
    defaultProperties() {
        return ['id', 'totalParticipantCount'];
    }
    constructor() {
        super();
        _TeamsMeetingAttendancereportListCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsMeetingAttendancereportListCommand_instances, "m", _TeamsMeetingAttendancereportListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsMeetingAttendancereportListCommand_instances, "m", _TeamsMeetingAttendancereportListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsMeetingAttendancereportListCommand_instances, "m", _TeamsMeetingAttendancereportListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken);
        if (isAppOnlyAccessToken && !args.options.userId && !args.options.userName && !args.options.email) {
            this.handleError(`The option 'userId', 'userName' or 'email' is required when retrieving meeting attendance report using app only permissions`);
        }
        else if (!isAppOnlyAccessToken && (args.options.userId || args.options.userName || args.options.email)) {
            this.handleError(`The options 'userId', 'userName' and 'email' cannot be used when retrieving meeting attendance reports using delegated permissions`);
        }
        try {
            if (this.verbose) {
                await logger.logToStderr(`Retrieving attendance report for ${isAppOnlyAccessToken ? 'specific user' : 'currently logged in user'}`);
            }
            let requestUrl = `${this.resource}/v1.0/`;
            if (isAppOnlyAccessToken) {
                requestUrl += 'users/';
                if (args.options.userId) {
                    requestUrl += args.options.userId;
                }
                else {
                    const userId = await this.getUserId(args.options.userName, args.options.email);
                    requestUrl += userId;
                }
            }
            else {
                requestUrl += `me`;
            }
            requestUrl += `/onlineMeetings/${args.options.meetingId}/attendanceReports`;
            const res = await odata.getAllItems(requestUrl);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getUserId(userName, email) {
        const options = {
            email: email,
            userName: userName,
            output: 'json',
            debug: this.debug,
            verbose: this.verbose
        };
        const output = await cli.executeCommandWithOutput(entraUserGetCommand, { options: { ...options, _: [] } });
        const getUserOutput = JSON.parse(output.stdout);
        return getUserOutput.id;
    }
}
_TeamsMeetingAttendancereportListCommand_instances = new WeakSet(), _TeamsMeetingAttendancereportListCommand_initTelemetry = function _TeamsMeetingAttendancereportListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            email: typeof args.options.email !== 'undefined'
        });
    });
}, _TeamsMeetingAttendancereportListCommand_initOptions = function _TeamsMeetingAttendancereportListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --userId [userId]'
    }, {
        option: '-n, --userName [userName]'
    }, {
        option: '--email [email]'
    }, {
        option: '-m, --meetingId <meetingId>'
    });
}, _TeamsMeetingAttendancereportListCommand_initValidators = function _TeamsMeetingAttendancereportListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid Guid`;
        }
        return true;
    });
};
export default new TeamsMeetingAttendancereportListCommand();
//# sourceMappingURL=meeting-attendancereport-list.js.map