var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsAppUpdateCommand_instances, _TeamsAppUpdateCommand_initTelemetry, _TeamsAppUpdateCommand_initOptions, _TeamsAppUpdateCommand_initValidators, _TeamsAppUpdateCommand_initOptionSets;
import fs from 'fs';
import path from 'path';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';
class TeamsAppUpdateCommand extends GraphCommand {
    get name() {
        return commands.APP_UPDATE;
    }
    get description() {
        return 'Updates Teams app in the organization\'s app catalog';
    }
    constructor() {
        super();
        _TeamsAppUpdateCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsAppUpdateCommand_instances, "m", _TeamsAppUpdateCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsAppUpdateCommand_instances, "m", _TeamsAppUpdateCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsAppUpdateCommand_instances, "m", _TeamsAppUpdateCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsAppUpdateCommand_instances, "m", _TeamsAppUpdateCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const { filePath } = args.options;
        try {
            const appId = await this.getAppId(args.options);
            const fullPath = path.resolve(filePath);
            if (this.verbose) {
                await logger.logToStderr(`Updating app with id '${appId}' and file '${fullPath}' in the app catalog...`);
            }
            const requestOptions = {
                url: `${this.resource}/v1.0/appCatalogs/teamsApps/${appId}`,
                headers: {
                    "content-type": "application/zip"
                },
                data: fs.readFileSync(fullPath)
            };
            await request.put(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
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
_TeamsAppUpdateCommand_instances = new WeakSet(), _TeamsAppUpdateCommand_initTelemetry = function _TeamsAppUpdateCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined'
        });
    });
}, _TeamsAppUpdateCommand_initOptions = function _TeamsAppUpdateCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-n, --name [name]'
    }, {
        option: '-p, --filePath <filePath>'
    });
}, _TeamsAppUpdateCommand_initValidators = function _TeamsAppUpdateCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        const fullPath = path.resolve(args.options.filePath);
        if (!fs.existsSync(fullPath)) {
            return `File '${fullPath}' not found`;
        }
        if (fs.lstatSync(fullPath).isDirectory()) {
            return `Path '${fullPath}' points to a directory`;
        }
        return true;
    });
}, _TeamsAppUpdateCommand_initOptionSets = function _TeamsAppUpdateCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] });
};
export default new TeamsAppUpdateCommand();
//# sourceMappingURL=app-update.js.map