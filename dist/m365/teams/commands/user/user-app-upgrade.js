var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsUserAppUpgradeCommand_instances, _TeamsUserAppUpgradeCommand_initTelemetry, _TeamsUserAppUpgradeCommand_initOptions, _TeamsUserAppUpgradeCommand_initValidators, _TeamsUserAppUpgradeCommand_initOptionSets, _TeamsUserAppUpgradeCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsUserAppUpgradeCommand extends GraphCommand {
    get name() {
        return commands.USER_APP_UPGRADE;
    }
    get description() {
        return 'Upgrade an app in the personal scope of the specified user';
    }
    constructor() {
        super();
        _TeamsUserAppUpgradeCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsUserAppUpgradeCommand_instances, "m", _TeamsUserAppUpgradeCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsUserAppUpgradeCommand_instances, "m", _TeamsUserAppUpgradeCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsUserAppUpgradeCommand_instances, "m", _TeamsUserAppUpgradeCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsUserAppUpgradeCommand_instances, "m", _TeamsUserAppUpgradeCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _TeamsUserAppUpgradeCommand_instances, "m", _TeamsUserAppUpgradeCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Upgrading app ${args.options.id || args.options.name} for user ${args.options.userId || args.options.userName}`);
            }
            const installedAppId = await this.getInstalledAppId(args.options, logger);
            const requestOptions = {
                url: `${this.resource}/v1.0/users/${formatting.encodeQueryParameter(args.options.userId || args.options.userName)}/teamwork/installedApps/${installedAppId}/upgrade`,
                headers: {
                    'accept': 'application/json;odata.metadata=none'
                }
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getInstalledAppId(options, logger) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving app ID`);
        }
        if (options.id) {
            return options.id;
        }
        const installedApps = await odata.getAllItems(`${this.resource}/v1.0/users/${formatting.encodeQueryParameter(options.userId || options.userName)}/teamwork/installedApps?$expand=teamsAppDefinition&$filter=teamsAppDefinition/displayName eq '${formatting.encodeQueryParameter(options.name)}'&$select=id`);
        if (installedApps.length === 1) {
            return installedApps[0].id;
        }
        if (installedApps.length === 0) {
            throw `The specified Teams app ${options.name} does not exist or is not installed for the user`;
        }
        const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', installedApps);
        const result = (await cli.handleMultipleResultsFound(`Multiple installed Teams apps with name '${options.name}' found.`, resultAsKeyValuePair));
        return result.id;
    }
}
_TeamsUserAppUpgradeCommand_instances = new WeakSet(), _TeamsUserAppUpgradeCommand_initTelemetry = function _TeamsUserAppUpgradeCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined'
        });
    });
}, _TeamsUserAppUpgradeCommand_initOptions = function _TeamsUserAppUpgradeCommand_initOptions() {
    this.options.unshift({
        option: '--id [id]'
    }, {
        option: '--name [name]'
    }, {
        option: '--userId [userId]'
    }, {
        option: '--userName [userName]'
    });
}, _TeamsUserAppUpgradeCommand_initValidators = function _TeamsUserAppUpgradeCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid userName`;
        }
        return true;
    });
}, _TeamsUserAppUpgradeCommand_initOptionSets = function _TeamsUserAppUpgradeCommand_initOptionSets() {
    this.optionSets.push({
        options: ['id', 'name']
    }, {
        options: ['userId', 'userName']
    });
}, _TeamsUserAppUpgradeCommand_initTypes = function _TeamsUserAppUpgradeCommand_initTypes() {
    this.types.string.push('id', 'name', 'userId', 'userName');
};
export default new TeamsUserAppUpgradeCommand();
//# sourceMappingURL=user-app-upgrade.js.map