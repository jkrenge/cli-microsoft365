var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsUserAppListCommand_instances, _TeamsUserAppListCommand_initTelemetry, _TeamsUserAppListCommand_initOptions, _TeamsUserAppListCommand_initValidators, _TeamsUserAppListCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsUserAppListCommand extends GraphCommand {
    get name() {
        return commands.USER_APP_LIST;
    }
    get description() {
        return 'List the apps installed in the personal scope of the specified user';
    }
    constructor() {
        super();
        _TeamsUserAppListCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsUserAppListCommand_instances, "m", _TeamsUserAppListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsUserAppListCommand_instances, "m", _TeamsUserAppListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsUserAppListCommand_instances, "m", _TeamsUserAppListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsUserAppListCommand_instances, "m", _TeamsUserAppListCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const userId = (await this.getUserId(args)).value;
            const endpoint = `${this.resource}/v1.0/users/${formatting.encodeQueryParameter(userId)}/teamwork/installedApps?$expand=teamsAppDefinition,teamsApp`;
            const items = await odata.getAllItems(endpoint);
            items.forEach(i => {
                const userAppId = Buffer.from(i.id, 'base64').toString('ascii');
                const appId = userAppId.substring(userAppId.indexOf("##") + 2, userAppId.length);
                i.appId = appId;
            });
            if (!cli.shouldTrimOutput(args.options.output)) {
                await logger.log(items);
            }
            else {
                await logger.log(items.map(i => {
                    return {
                        id: i.id,
                        appId: i.appId,
                        displayName: i.teamsAppDefinition.displayName,
                        version: i.teamsAppDefinition.version
                    };
                }));
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getUserId(args) {
        if (args.options.userId) {
            return { value: args.options.userId };
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/users/${formatting.encodeQueryParameter(args.options.userName)}/id`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        return request.get(requestOptions);
    }
}
_TeamsUserAppListCommand_instances = new WeakSet(), _TeamsUserAppListCommand_initTelemetry = function _TeamsUserAppListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined'
        });
    });
}, _TeamsUserAppListCommand_initOptions = function _TeamsUserAppListCommand_initOptions() {
    this.options.unshift({
        option: '--userId [userId]'
    }, {
        option: '--userName [userName]'
    });
}, _TeamsUserAppListCommand_initValidators = function _TeamsUserAppListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid userName`;
        }
        return true;
    });
}, _TeamsUserAppListCommand_initOptionSets = function _TeamsUserAppListCommand_initOptionSets() {
    this.optionSets.push({ options: ['userId', 'userName'] });
};
export default new TeamsUserAppListCommand();
//# sourceMappingURL=user-app-list.js.map