var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraAppGetCommand_instances, _EntraAppGetCommand_initTelemetry, _EntraAppGetCommand_initOptions, _EntraAppGetCommand_initValidators, _EntraAppGetCommand_initOptionSets;
import fs from 'fs';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { entraApp } from '../../../../utils/entraApp.js';
class EntraAppGetCommand extends GraphCommand {
    get name() {
        return commands.APP_GET;
    }
    get description() {
        return 'Gets an Entra app registration';
    }
    constructor() {
        super();
        _EntraAppGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraAppGetCommand_instances, "m", _EntraAppGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraAppGetCommand_instances, "m", _EntraAppGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraAppGetCommand_instances, "m", _EntraAppGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraAppGetCommand_instances, "m", _EntraAppGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const appInfo = await this.getAppInfo(args, logger);
            const res = await this.saveAppInfo(args, appInfo, logger);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getAppInfo(args, logger) {
        const { objectId, appId, name } = args.options;
        const properties = args.options.properties?.split(',');
        if (this.verbose) {
            await logger.logToStderr(`Retrieving information about Microsoft Entra app ${objectId ? objectId : (appId ? appId : name)}...`);
        }
        if (objectId) {
            return await entraApp.getAppRegistrationByObjectId(objectId, properties);
        }
        else if (appId) {
            return await entraApp.getAppRegistrationByAppId(appId, properties);
        }
        else {
            return await entraApp.getAppRegistrationByAppName(name, properties);
        }
    }
    async saveAppInfo(args, appInfo, logger) {
        if (!args.options.save) {
            return appInfo;
        }
        const filePath = '.m365rc.json';
        if (this.verbose) {
            await logger.logToStderr(`Saving Microsoft Entra app registration information to the ${filePath} file...`);
        }
        let m365rc = {};
        if (fs.existsSync(filePath)) {
            if (this.debug) {
                await logger.logToStderr(`Reading existing ${filePath}...`);
            }
            try {
                const fileContents = fs.readFileSync(filePath, 'utf8');
                if (fileContents) {
                    m365rc = JSON.parse(fileContents);
                }
            }
            catch (e) {
                await logger.logToStderr(`Error reading ${filePath}: ${e}. Please add app info to ${filePath} manually.`);
                return Promise.resolve(appInfo);
            }
        }
        if (!m365rc.apps) {
            m365rc.apps = [];
        }
        if (!m365rc.apps.some(a => a.appId === appInfo.appId)) {
            m365rc.apps.push({
                appId: appInfo.appId,
                name: appInfo.displayName
            });
            try {
                fs.writeFileSync(filePath, JSON.stringify(m365rc, null, 2));
            }
            catch (e) {
                await logger.logToStderr(`Error writing ${filePath}: ${e}. Please add app info to ${filePath} manually.`);
            }
        }
        return appInfo;
    }
}
_EntraAppGetCommand_instances = new WeakSet(), _EntraAppGetCommand_initTelemetry = function _EntraAppGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            appId: typeof args.options.appId !== 'undefined',
            objectId: typeof args.options.objectId !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            properties: typeof args.options.properties !== 'undefined'
        });
    });
}, _EntraAppGetCommand_initOptions = function _EntraAppGetCommand_initOptions() {
    this.options.unshift({ option: '--appId [appId]' }, { option: '--objectId [objectId]' }, { option: '--name [name]' }, { option: '--save' }, { option: '-p, --properties [properties]' });
}, _EntraAppGetCommand_initValidators = function _EntraAppGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.appId && !validation.isValidGuid(args.options.appId)) {
            return `${args.options.appId} is not a valid GUID`;
        }
        if (args.options.objectId && !validation.isValidGuid(args.options.objectId)) {
            return `${args.options.objectId} is not a valid GUID`;
        }
        return true;
    });
}, _EntraAppGetCommand_initOptionSets = function _EntraAppGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['appId', 'objectId', 'name'] });
};
export default new EntraAppGetCommand();
//# sourceMappingURL=app-get.js.map