var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsMeetingTranscriptGetCommand_instances, _TeamsMeetingTranscriptGetCommand_initTelemetry, _TeamsMeetingTranscriptGetCommand_initOptions, _TeamsMeetingTranscriptGetCommand_initValidators, _TeamsMeetingTranscriptGetCommand_initOptionSets;
import auth from '../../../../Auth.js';
import request from '../../../../request.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import fs from 'fs';
import path from 'path';
class TeamsMeetingTranscriptGetCommand extends GraphCommand {
    get name() {
        return commands.MEETING_TRANSCRIPT_GET;
    }
    get description() {
        return 'Downloads a transcript for a given meeting';
    }
    constructor() {
        super();
        _TeamsMeetingTranscriptGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsMeetingTranscriptGetCommand_instances, "m", _TeamsMeetingTranscriptGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsMeetingTranscriptGetCommand_instances, "m", _TeamsMeetingTranscriptGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsMeetingTranscriptGetCommand_instances, "m", _TeamsMeetingTranscriptGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsMeetingTranscriptGetCommand_instances, "m", _TeamsMeetingTranscriptGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken);
            if (this.verbose) {
                await logger.logToStderr(`Retrieving transcript for the given meeting...`);
            }
            let requestUrl = `${this.resource}/beta/`;
            if (isAppOnlyAccessToken) {
                if (!args.options.userId && !args.options.userName && !args.options.email) {
                    throw `The option 'userId', 'userName' or 'email' is required when retrieving meeting transcript using app only permissions`;
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
                    throw `The options 'userId', 'userName', and 'email' cannot be used while retrieving meeting transcript using delegated permissions`;
                }
                requestUrl += `me`;
            }
            requestUrl += `/onlineMeetings/${args.options.meetingId}/transcripts/${args.options.id}`;
            if (args.options.outputFile) {
                requestUrl += '/content?$format=text/vtt';
            }
            const requestOptions = {
                url: requestUrl,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: args.options.outputFile ? 'stream' : 'json'
            };
            const meetingTranscript = await request.get(requestOptions);
            if (meetingTranscript) {
                if (args.options.outputFile) {
                    // Not possible to use async/await for this promise
                    await new Promise((resolve, reject) => {
                        const writer = fs.createWriteStream(args.options.outputFile);
                        meetingTranscript.data.pipe(writer);
                        writer.on('error', err => {
                            reject(err);
                        });
                        writer.on('close', async () => {
                            const filePath = args.options.outputFile;
                            if (this.verbose) {
                                await logger.logToStderr(`File saved to path ${filePath}`);
                            }
                            return resolve();
                        });
                    });
                }
                else {
                    await logger.log(meetingTranscript);
                }
            }
            else {
                throw `The specified meeting transcript was not found`;
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_TeamsMeetingTranscriptGetCommand_instances = new WeakSet(), _TeamsMeetingTranscriptGetCommand_initTelemetry = function _TeamsMeetingTranscriptGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            email: typeof args.options.email !== 'undefined',
            outputFile: typeof args.options.outputFile !== 'undefined'
        });
    });
}, _TeamsMeetingTranscriptGetCommand_initOptions = function _TeamsMeetingTranscriptGetCommand_initOptions() {
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
    }, {
        option: '-f, --outputFile [outputFile]'
    });
}, _TeamsMeetingTranscriptGetCommand_initValidators = function _TeamsMeetingTranscriptGetCommand_initValidators() {
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
        if (args.options.outputFile && !fs.existsSync(path.dirname(args.options.outputFile))) {
            return 'Specified path where to save the file does not exits';
        }
        return true;
    });
}, _TeamsMeetingTranscriptGetCommand_initOptionSets = function _TeamsMeetingTranscriptGetCommand_initOptionSets() {
    this.optionSets.push({
        options: ['userId', 'userName', 'email'],
        runsWhen: (args) => args.options.userId || args.options.userName || args.options.email
    });
};
export default new TeamsMeetingTranscriptGetCommand();
//# sourceMappingURL=meeting-transcript-get.js.map