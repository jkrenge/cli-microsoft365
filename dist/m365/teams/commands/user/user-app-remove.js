var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsUserAppRemoveCommand_instances, _TeamsUserAppRemoveCommand_initTelemetry, _TeamsUserAppRemoveCommand_initOptions, _TeamsUserAppRemoveCommand_initValidators, _TeamsUserAppRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsUserAppRemoveCommand extends GraphCommand {
    get name() {
        return commands.USER_APP_REMOVE;
    }
    get description() {
        return 'Uninstall an app from the personal scope of the specified user.';
    }
    constructor() {
        super();
        _TeamsUserAppRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsUserAppRemoveCommand_instances, "m", _TeamsUserAppRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsUserAppRemoveCommand_instances, "m", _TeamsUserAppRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsUserAppRemoveCommand_instances, "m", _TeamsUserAppRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsUserAppRemoveCommand_instances, "m", _TeamsUserAppRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const removeApp = async () => {
            const appId = await this.getAppId(args);
            // validation ensures that here we have either userId or userName
            const userId = (args.options.userId ?? args.options.userName);
            const endpoint = `${this.resource}/v1.0`;
            if (this.verbose) {
                await logger.logToStderr(`Removing app with ID ${args.options.id} for user ${args.options.userId}`);
            }
            const requestOptions = {
                url: `${endpoint}/users/${formatting.encodeQueryParameter(userId)}/teamwork/installedApps/${appId}`,
                headers: {
                    'accept': 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            try {
                await request.delete(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeApp();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the app with id ${args.options.id} for user ${args.options.userId ?? args.options.userName}?` });
            if (result) {
                await removeApp();
            }
        }
    }
    async getAppId(args) {
        if (args.options.id) {
            return args.options.id;
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/users/${args.options.userId}/teamwork/installedApps?$expand=teamsAppDefinition&$filter=teamsAppDefinition/displayName eq '${formatting.encodeQueryParameter(args.options.name)}'`,
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
_TeamsUserAppRemoveCommand_instances = new WeakSet(), _TeamsUserAppRemoveCommand_initTelemetry = function _TeamsUserAppRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            force: (!!args.options.force).toString()
        });
    });
}, _TeamsUserAppRemoveCommand_initOptions = function _TeamsUserAppRemoveCommand_initOptions() {
    this.options.unshift({
        option: '--id [id]'
    }, {
        option: '--name [name]'
    }, {
        option: '--userId [userId]'
    }, {
        option: '--userName [userName]'
    }, {
        option: '-f, --force'
    });
}, _TeamsUserAppRemoveCommand_initValidators = function _TeamsUserAppRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid userName`;
        }
        return true;
    });
}, _TeamsUserAppRemoveCommand_initOptionSets = function _TeamsUserAppRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] });
    this.optionSets.push({ options: ['userId', 'userName'] });
};
export default new TeamsUserAppRemoveCommand();
//# sourceMappingURL=user-app-remove.js.map