var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsMeetingListCommand_instances, _TeamsMeetingListCommand_initTelemetry, _TeamsMeetingListCommand_initOptions, _TeamsMeetingListCommand_initValidators;
import auth from '../../../../Auth.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from "../../../base/GraphCommand.js";
import commands from '../../commands.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
class TeamsMeetingListCommand extends GraphCommand {
    get name() {
        return commands.MEETING_LIST;
    }
    get description() {
        return 'Retrieve all online meetings for a given user or shared mailbox';
    }
    defaultProperties() {
        return ['subject', 'startDateTime', 'endDateTime'];
    }
    constructor() {
        super();
        _TeamsMeetingListCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsMeetingListCommand_instances, "m", _TeamsMeetingListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsMeetingListCommand_instances, "m", _TeamsMeetingListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsMeetingListCommand_instances, "m", _TeamsMeetingListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken);
            if (isAppOnlyAccessToken && !args.options.userId && !args.options.userName && !args.options.email) {
                throw `The option 'userId', 'userName' or 'email' is required when retrieving meetings using app only permissions`;
            }
            else if (!isAppOnlyAccessToken && (args.options.userId || args.options.userName || args.options.email)) {
                throw `The options 'userId', 'userName' and 'email' cannot be used when retrieving meetings using delegated permissions`;
            }
            if (this.verbose) {
                await logger.logToStderr(`Retrieving meetings for user: ${args.options.userId || args.options.userName || args.options.email || accessToken.getUserNameFromAccessToken(auth.connection.accessTokens[this.resource].accessToken)}...`);
            }
            const graphBaseUrl = await this.getGraphBaseUrl(args.options);
            const meetingUrls = await this.getMeetingJoinUrls(graphBaseUrl, args.options);
            const meetings = await this.getTeamsMeetings(logger, graphBaseUrl, meetingUrls);
            await logger.log(meetings);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    /**
     * Get the first part of the Graph API URL that contains the user information.
     */
    async getGraphBaseUrl(options) {
        let requestUrl = `${this.resource}/v1.0/`;
        if (options.userId || options.userName) {
            requestUrl += `users/${options.userId || options.userName}`;
        }
        else if (options.email) {
            const userId = await entraUser.getUserIdByEmail(options.email);
            requestUrl += `users/${userId}`;
        }
        else {
            requestUrl += 'me';
        }
        return requestUrl;
    }
    /**
     * Gets the meeting join urls for the specified user using calendar events.
     */
    async getMeetingJoinUrls(graphBaseUrl, options) {
        let requestUrl = graphBaseUrl;
        requestUrl += `/events?$filter=start/dateTime ge '${options.startDateTime}'`;
        if (options.endDateTime) {
            requestUrl += ` and end/dateTime lt '${options.endDateTime}'`;
        }
        if (options.isOrganizer) {
            requestUrl += ' and isOrganizer eq true';
        }
        requestUrl += '&$select=onlineMeeting';
        const items = await odata.getAllItems(requestUrl);
        const result = items.filter(i => i.onlineMeeting).map(i => i.onlineMeeting.joinUrl);
        return result;
    }
    async getTeamsMeetings(logger, graphBaseUrl, meetingUrls) {
        const graphRelativeUrl = graphBaseUrl.replace(`${this.resource}/v1.0/`, '');
        let result = [];
        for (let i = 0; i < meetingUrls.length; i += 20) {
            if (this.verbose) {
                await logger.logToStderr(`Retrieving meetings ${i + 1}-${Math.min(i + 20, meetingUrls.length)}...`);
            }
            const batch = meetingUrls.slice(i, i + 20);
            const requestOptions = {
                url: `${this.resource}/v1.0/$batch`,
                headers: {
                    accept: 'application/json',
                    'content-type': 'application/json'
                },
                responseType: 'json',
                data: {
                    requests: batch.map((url, index) => ({
                        id: i + index,
                        method: 'GET',
                        url: `${graphRelativeUrl}/onlineMeetings?$filter=joinWebUrl eq '${formatting.encodeQueryParameter(url)}'`
                    }))
                }
            };
            const requestResponse = await request.post(requestOptions);
            for (const response of requestResponse.responses) {
                if (response.status === 200) {
                    result.push(response.body.value[0]);
                }
                else {
                    // Encountered errors where message was empty resulting in [object Object] error messages
                    if (!response.body.error.message) {
                        throw response.body.error.code;
                    }
                    throw response.body;
                }
            }
        }
        // Sort all meetings by start date
        result = result.sort((a, b) => a.startDateTime < b.startDateTime ? -1 : a.startDateTime > b.startDateTime ? 1 : 0);
        return result;
    }
}
_TeamsMeetingListCommand_instances = new WeakSet(), _TeamsMeetingListCommand_initTelemetry = function _TeamsMeetingListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            email: typeof args.options.email !== 'undefined',
            endDateTime: typeof args.options.endDateTime !== 'undefined',
            isOrganizer: !!args.options.isOrganizer
        });
    });
}, _TeamsMeetingListCommand_initOptions = function _TeamsMeetingListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --userId [userId]'
    }, {
        option: '-n, --userName [userName]'
    }, {
        option: '--email [email]'
    }, {
        option: '--startDateTime <startDateTime>'
    }, {
        option: '--endDateTime [endDateTime]'
    }, {
        option: '--isOrganizer'
    });
}, _TeamsMeetingListCommand_initValidators = function _TeamsMeetingListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidISODateTime(args.options.startDateTime)) {
            return `'${args.options.startDateTime}' is not a valid ISO date string for startDateTime.`;
        }
        if (args.options.endDateTime && !validation.isValidISODateTime(args.options.endDateTime)) {
            return `'${args.options.startDateTime}' is not a valid ISO date string for endDateTime.`;
        }
        if (args.options.startDateTime && args.options.endDateTime && args.options.startDateTime > args.options.endDateTime) {
            return 'startDateTime value must be before endDateTime.';
        }
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID for userId.`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `'${args.options.userName}' is not a valid UPN for userName.`;
        }
        if (args.options.email && !validation.isValidUserPrincipalName(args.options.email)) {
            return `'${args.options.email}' is not a valid UPN for email.`;
        }
        return true;
    });
};
export default new TeamsMeetingListCommand();
//# sourceMappingURL=meeting-list.js.map