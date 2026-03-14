var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsTabAddCommand_instances, _TeamsTabAddCommand_initTelemetry, _TeamsTabAddCommand_initOptions, _TeamsTabAddCommand_initValidators;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsTabAddCommand extends GraphCommand {
    get name() {
        return commands.TAB_ADD;
    }
    allowUnknownOptions() {
        return true;
    }
    get description() {
        return 'Add a tab to the specified channel';
    }
    constructor() {
        super();
        _TeamsTabAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsTabAddCommand_instances, "m", _TeamsTabAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsTabAddCommand_instances, "m", _TeamsTabAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsTabAddCommand_instances, "m", _TeamsTabAddCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const data = this.mapRequestBody(args.options);
        const requestOptions = {
            url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(args.options.teamId)}/channels/${args.options.channelId}/tabs`,
            headers: {
                accept: 'application/json;odata.metadata=none',
                'content-type': 'application/json;odata=nometadata'
            },
            data: data,
            responseType: 'json'
        };
        try {
            const res = await request.post(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    mapRequestBody(options) {
        const requestBody = {};
        requestBody['configuration'] = {};
        const excludeOptions = [
            'debug',
            'verbose',
            'teamId',
            'channelId',
            'appId',
            'appName',
            'entityId',
            'contentUrl',
            'removeUrl',
            'websiteUrl',
            'output'
        ];
        if (options.appName) {
            requestBody.displayName = options.appName;
        }
        if (options.appId) {
            requestBody['teamsApp@odata.bind'] = `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/${options.appId}`;
        }
        if (options.contentUrl) {
            requestBody.configuration.contentUrl = options.contentUrl;
        }
        if (options.entityId) {
            requestBody.configuration.entityId = options.entityId;
        }
        if (options.removeUrl) {
            requestBody.configuration.removeUrl = options.removeUrl;
        }
        if (options.websiteUrl) {
            requestBody.configuration.websiteUrl = options.websiteUrl;
        }
        Object.keys(options).forEach(key => {
            if (excludeOptions.indexOf(key) === -1) {
                requestBody.configuration[key] = `${options[key]}`;
            }
        });
        return requestBody;
    }
}
_TeamsTabAddCommand_instances = new WeakSet(), _TeamsTabAddCommand_initTelemetry = function _TeamsTabAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            entityId: typeof args.options.entityId !== 'undefined',
            removeUrl: typeof args.options.removeUrl !== 'undefined',
            websiteUrl: typeof args.options.websiteUrl !== 'undefined'
        });
    });
}, _TeamsTabAddCommand_initOptions = function _TeamsTabAddCommand_initOptions() {
    this.options.unshift({
        option: '-i, --teamId <teamId>'
    }, {
        option: '-c, --channelId <channelId>'
    }, {
        option: '--appId <appId>'
    }, {
        option: '--appName <appName>'
    }, {
        option: '--contentUrl <contentUrl>'
    }, {
        option: '--entityId [entityId]'
    }, {
        option: '--removeUrl [removeUrl]'
    }, {
        option: '--websiteUrl [websiteUrl]'
    });
}, _TeamsTabAddCommand_initValidators = function _TeamsTabAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.teamId)) {
            return `${args.options.teamId} is not a valid GUID`;
        }
        if (!validation.isValidTeamsChannelId(args.options.channelId)) {
            return `${args.options.channelId} is not a valid Teams ChannelId`;
        }
        return true;
    });
};
export default new TeamsTabAddCommand();
//# sourceMappingURL=tab-add.js.map