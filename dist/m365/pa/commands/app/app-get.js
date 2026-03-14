var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PaAppGetCommand_instances, _PaAppGetCommand_initTelemetry, _PaAppGetCommand_initOptions, _PaAppGetCommand_initValidators, _PaAppGetCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import PowerAppsCommand from '../../../base/PowerAppsCommand.js';
import commands from '../../commands.js';
import paAppListCommand from '../app/app-list.js';
class PaAppGetCommand extends PowerAppsCommand {
    get name() {
        return commands.APP_GET;
    }
    get description() {
        return 'Gets information about the specified Microsoft Power App';
    }
    constructor() {
        super();
        _PaAppGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _PaAppGetCommand_instances, "m", _PaAppGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PaAppGetCommand_instances, "m", _PaAppGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PaAppGetCommand_instances, "m", _PaAppGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PaAppGetCommand_instances, "m", _PaAppGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (args.options.name) {
                let endpoint = `${this.resource}/providers/Microsoft.PowerApps`;
                if (args.options.asAdmin) {
                    endpoint += `/scopes/admin/environments/${formatting.encodeQueryParameter(args.options.environmentName)}`;
                }
                endpoint += `/apps/${formatting.encodeQueryParameter(args.options.name)}?api-version=2016-11-01`;
                const requestOptions = {
                    url: endpoint,
                    headers: {
                        accept: 'application/json'
                    },
                    responseType: 'json'
                };
                if (this.verbose) {
                    await logger.logToStderr(`Retrieving information about Microsoft Power App with name '${args.options.name}'...`);
                }
                const res = await request.get(requestOptions);
                await logger.log(this.setProperties(res));
            }
            else {
                if (this.verbose) {
                    await logger.logToStderr(`Retrieving information about Microsoft Power App with displayName '${args.options.displayName}'...`);
                }
                const getAppsOutput = await this.getApps(args, logger);
                if (getAppsOutput.stdout && JSON.parse(getAppsOutput.stdout).length > 0) {
                    const allApps = JSON.parse(getAppsOutput.stdout);
                    const app = allApps.find((a) => {
                        return a.properties.displayName.toLowerCase() === `${args.options.displayName}`.toLowerCase();
                    });
                    if (app) {
                        await logger.log(this.setProperties(app));
                    }
                    else {
                        throw `No app found with displayName '${args.options.displayName}'.`;
                    }
                }
                else {
                    throw 'No apps found.';
                }
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getApps(args, logger) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving all apps...`);
        }
        const options = {
            output: 'json',
            debug: this.debug,
            verbose: this.verbose,
            environmentName: args.options.environmentName,
            asAdmin: args.options.asAdmin
        };
        return await cli.executeCommandWithOutput(paAppListCommand, { options: { ...options, _: [] } });
    }
    setProperties(app) {
        app.displayName = app.properties.displayName;
        app.description = app.properties.description || '';
        app.appVersion = app.properties.appVersion;
        app.owner = app.properties.owner.email || '';
        return app;
    }
}
_PaAppGetCommand_instances = new WeakSet(), _PaAppGetCommand_initTelemetry = function _PaAppGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            name: typeof args.options.name !== 'undefined',
            displayName: typeof args.options.displayName !== 'undefined',
            asAdmin: !!args.options.asAdmin,
            environmentName: typeof args.options.environmentName !== 'undefined'
        });
    });
}, _PaAppGetCommand_initOptions = function _PaAppGetCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name [name]'
    }, {
        option: '-d, --displayName [displayName]'
    }, {
        option: '-e, --environmentName [environmentName]'
    }, {
        option: '--asAdmin'
    });
}, _PaAppGetCommand_initValidators = function _PaAppGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.name && !validation.isValidGuid(args.options.name)) {
            return `${args.options.name} is not a valid GUID`;
        }
        if (args.options.asAdmin && !args.options.environmentName) {
            return 'When specifying the asAdmin option, the environment option is required as well.';
        }
        if (args.options.environmentName && !args.options.asAdmin) {
            return 'When specifying the environment option, the asAdmin option is required as well.';
        }
        return true;
    });
}, _PaAppGetCommand_initOptionSets = function _PaAppGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['name', 'displayName'] });
};
export default new PaAppGetCommand();
//# sourceMappingURL=app-get.js.map