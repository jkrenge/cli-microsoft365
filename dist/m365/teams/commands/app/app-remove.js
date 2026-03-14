var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsAppRemoveCommand_instances, _TeamsAppRemoveCommand_initTelemetry, _TeamsAppRemoveCommand_initOptions, _TeamsAppRemoveCommand_initValidators, _TeamsAppRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsAppRemoveCommand extends GraphCommand {
    get name() {
        return commands.APP_REMOVE;
    }
    get description() {
        return 'Removes a Teams app from the organization\'s app catalog';
    }
    constructor() {
        super();
        _TeamsAppRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsAppRemoveCommand_instances, "m", _TeamsAppRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsAppRemoveCommand_instances, "m", _TeamsAppRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsAppRemoveCommand_instances, "m", _TeamsAppRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsAppRemoveCommand_instances, "m", _TeamsAppRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const removeApp = async () => {
            try {
                const appId = await this.getAppId(args.options, logger);
                if (this.verbose) {
                    await logger.logToStderr(`Removing app with ID ${appId}`);
                }
                const requestOptions = {
                    url: `${this.resource}/v1.0/appCatalogs/teamsApps/${appId}`,
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    }
                };
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
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the Teams app ${args.options.id || args.options.name} from the app catalog?` });
            if (result) {
                await removeApp();
            }
        }
    }
    async getAppId(options, logger) {
        if (options.id) {
            return options.id;
        }
        if (this.verbose) {
            await logger.logToStderr(`Retrieving app Id...`);
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/appCatalogs/teamsApps?$filter=displayName eq '${formatting.encodeQueryParameter(options.name)}'&$select=id`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        const app = response.value[0];
        if (!app) {
            throw `The specified Teams app does not exist`;
        }
        if (response.value.length > 1) {
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', response.value);
            const result = await cli.handleMultipleResultsFound(`Multiple Teams apps with name '${options.name}' found.`, resultAsKeyValuePair);
            return result.id;
        }
        return app.id;
    }
}
_TeamsAppRemoveCommand_instances = new WeakSet(), _TeamsAppRemoveCommand_initTelemetry = function _TeamsAppRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: (!(!args.options.force)).toString(),
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined'
        });
    });
}, _TeamsAppRemoveCommand_initOptions = function _TeamsAppRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-n, --name [name]'
    }, {
        option: '-f, --force'
    });
}, _TeamsAppRemoveCommand_initValidators = function _TeamsAppRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
}, _TeamsAppRemoveCommand_initOptionSets = function _TeamsAppRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] });
};
export default new TeamsAppRemoveCommand();
//# sourceMappingURL=app-remove.js.map