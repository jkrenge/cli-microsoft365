var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsUserAppAddCommand_instances, _TeamsUserAppAddCommand_initTelemetry, _TeamsUserAppAddCommand_initOptions, _TeamsUserAppAddCommand_initValidators, _TeamsUserAppAddCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsUserAppAddCommand extends GraphCommand {
    get name() {
        return commands.USER_APP_ADD;
    }
    get description() {
        return 'Install an app in the personal scope of the specified user';
    }
    constructor() {
        super();
        _TeamsUserAppAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsUserAppAddCommand_instances, "m", _TeamsUserAppAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsUserAppAddCommand_instances, "m", _TeamsUserAppAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsUserAppAddCommand_instances, "m", _TeamsUserAppAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsUserAppAddCommand_instances, "m", _TeamsUserAppAddCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const appId = await this.getAppId(args);
            const userId = (args.options.userId ?? args.options.userName);
            const endpoint = `${this.resource}/v1.0`;
            if (this.verbose) {
                await logger.logToStderr(`Adding app with ID ${appId} for user ${args.options.userId}`);
            }
            const requestOptions = {
                url: `${endpoint}/users/${formatting.encodeQueryParameter(userId)}/teamwork/installedApps`,
                headers: {
                    'content-type': 'application/json;odata=nometadata',
                    'accept': 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: {
                    'teamsApp@odata.bind': `${endpoint}/appCatalogs/teamsApps/${appId}`
                }
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getAppId(args) {
        if (args.options.id) {
            return args.options.id;
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/appCatalogs/teamsApps?$filter=displayName eq '${formatting.encodeQueryParameter(args.options.name)}'`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        if (response.value.length === 1) {
            return response.value[0].id;
        }
        if (response.value.length === 0) {
            throw `The specified Teams app does not exist`;
        }
        const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', response.value);
        const result = (await cli.handleMultipleResultsFound(`Multiple Teams apps with name '${args.options.name}' found.`, resultAsKeyValuePair));
        return result.id;
    }
}
_TeamsUserAppAddCommand_instances = new WeakSet(), _TeamsUserAppAddCommand_initTelemetry = function _TeamsUserAppAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined'
        });
    });
}, _TeamsUserAppAddCommand_initOptions = function _TeamsUserAppAddCommand_initOptions() {
    this.options.unshift({
        option: '--id [id]'
    }, {
        option: '--name [name]'
    }, {
        option: '--userId [userId]'
    }, {
        option: '--userName [userName]'
    });
}, _TeamsUserAppAddCommand_initValidators = function _TeamsUserAppAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid userName`;
        }
        return true;
    });
}, _TeamsUserAppAddCommand_initOptionSets = function _TeamsUserAppAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] });
    this.optionSets.push({ options: ['userId', 'userName'] });
};
export default new TeamsUserAppAddCommand();
//# sourceMappingURL=user-app-add.js.map