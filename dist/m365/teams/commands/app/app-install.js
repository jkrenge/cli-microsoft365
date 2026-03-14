var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsAppInstallCommand_instances, _TeamsAppInstallCommand_initTelemetry, _TeamsAppInstallCommand_initOptions, _TeamsAppInstallCommand_initValidators, _TeamsAppInstallCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsAppInstallCommand extends GraphCommand {
    get name() {
        return commands.APP_INSTALL;
    }
    get description() {
        return 'Installs a Microsoft Teams team app from the catalog in the specified team or for the specified user';
    }
    constructor() {
        super();
        _TeamsAppInstallCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsAppInstallCommand_instances, "m", _TeamsAppInstallCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsAppInstallCommand_instances, "m", _TeamsAppInstallCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsAppInstallCommand_instances, "m", _TeamsAppInstallCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsAppInstallCommand_instances, "m", _TeamsAppInstallCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            await this.validateUser(args, logger);
            const appId = await this.getAppId(args.options);
            let url = `${this.resource}/v1.0`;
            if (args.options.teamId) {
                url += `/teams/${formatting.encodeQueryParameter(args.options.teamId)}/installedApps`;
            }
            else {
                url += `/users/${formatting.encodeQueryParameter((args.options.userId ?? args.options.userName))}/teamwork/installedApps`;
            }
            const requestOptions = {
                url: url,
                headers: {
                    'content-type': 'application/json;odata=nometadata',
                    'accept': 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: {
                    'teamsApp@odata.bind': `${this.resource}/v1.0/appCatalogs/teamsApps/${appId}`
                }
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    // we need this method, because passing an invalid user ID to the API
    // won't cause an error
    async validateUser(args, logger) {
        if (!args.options.userId) {
            return true;
        }
        if (this.verbose) {
            await logger.logToStderr(`Checking if user ${args.options.userId} exists...`);
        }
        try {
            const res = await entraUser.getUpnByUserId(args.options.userId, logger, this.verbose);
            if (this.verbose) {
                await logger.logToStderr(res);
            }
            return true;
        }
        catch (err) {
            if (this.verbose) {
                await logger.logToStderr(err.stderr);
            }
            throw `User with ID ${args.options.userId} not found. Original error: ${err.error.message}`;
        }
    }
    async getAppId(options) {
        if (options.id) {
            return options.id;
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/appCatalogs/teamsApps?$filter=displayName eq '${formatting.encodeQueryParameter(options.name)}'`,
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
            const result = await cli.handleMultipleResultsFound(`Multiple Teams apps with name ${options.name} found.`, resultAsKeyValuePair);
            return result.id;
        }
        return app.id;
    }
}
_TeamsAppInstallCommand_instances = new WeakSet(), _TeamsAppInstallCommand_initTelemetry = function _TeamsAppInstallCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            teamId: typeof args.options.teamId !== 'undefined',
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined'
        });
    });
}, _TeamsAppInstallCommand_initOptions = function _TeamsAppInstallCommand_initOptions() {
    this.options.unshift({ option: '-i, --id [id]' }, { option: '-n, --name [name]' }, { option: '--teamId [teamId]' }, { option: '--userId [userId]' }, { option: '--userName [userName]' });
}, _TeamsAppInstallCommand_initValidators = function _TeamsAppInstallCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        if (args.options.teamId &&
            !validation.isValidGuid(args.options.teamId)) {
            return `${args.options.teamId} is not a valid GUID`;
        }
        if (args.options.userId &&
            !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID`;
        }
        return true;
    });
}, _TeamsAppInstallCommand_initOptionSets = function _TeamsAppInstallCommand_initOptionSets() {
    this.optionSets.push({ options: ['teamId', 'userId', 'userName'] });
    this.optionSets.push({ options: ['id', 'name'] });
};
export default new TeamsAppInstallCommand();
//# sourceMappingURL=app-install.js.map