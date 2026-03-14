var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraAppRemoveCommand_instances, _EntraAppRemoveCommand_initTelemetry, _EntraAppRemoveCommand_initOptions, _EntraAppRemoveCommand_initValidators, _EntraAppRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { entraApp } from '../../../../utils/entraApp.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraAppRemoveCommand extends GraphCommand {
    get name() {
        return commands.APP_REMOVE;
    }
    get description() {
        return 'Removes an Entra app registration';
    }
    constructor() {
        super();
        _EntraAppRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraAppRemoveCommand_instances, "m", _EntraAppRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraAppRemoveCommand_instances, "m", _EntraAppRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraAppRemoveCommand_instances, "m", _EntraAppRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraAppRemoveCommand_instances, "m", _EntraAppRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const deleteApp = async () => {
            try {
                const objectId = await this.getObjectId(args, logger);
                if (this.verbose) {
                    await logger.logToStderr(`Deleting Microsoft Entra app ${objectId}...`);
                }
                const requestOptions = {
                    url: `${this.resource}/v1.0/myorganization/applications/${objectId}`,
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    },
                    responseType: 'json'
                };
                await request.delete(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await deleteApp();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the app?` });
            if (result) {
                await deleteApp();
            }
        }
    }
    async getObjectId(args, logger) {
        if (args.options.objectId) {
            return args.options.objectId;
        }
        const { appId, name } = args.options;
        if (this.verbose) {
            await logger.logToStderr(`Retrieving information about Microsoft Entra app ${appId ? appId : name}...`);
        }
        if (appId) {
            const app = await entraApp.getAppRegistrationByAppId(appId, ['id']);
            return app.id;
        }
        else {
            const app = await entraApp.getAppRegistrationByAppName(name, ['id']);
            return app.id;
        }
    }
}
_EntraAppRemoveCommand_instances = new WeakSet(), _EntraAppRemoveCommand_initTelemetry = function _EntraAppRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            appId: typeof args.options.appId !== 'undefined',
            objectId: typeof args.options.objectId !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            force: (!(!args.options.force)).toString()
        });
    });
}, _EntraAppRemoveCommand_initOptions = function _EntraAppRemoveCommand_initOptions() {
    this.options.unshift({ option: '--appId [appId]' }, { option: '--objectId [objectId]' }, { option: '--name [name]' }, { option: '-f, --force' });
}, _EntraAppRemoveCommand_initValidators = function _EntraAppRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.appId && !validation.isValidGuid(args.options.appId)) {
            return `${args.options.appId} is not a valid GUID`;
        }
        if (args.options.objectId && !validation.isValidGuid(args.options.objectId)) {
            return `${args.options.objectId} is not a valid GUID`;
        }
        return true;
    });
}, _EntraAppRemoveCommand_initOptionSets = function _EntraAppRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['appId', 'objectId', 'name'] });
};
export default new EntraAppRemoveCommand();
//# sourceMappingURL=app-remove.js.map