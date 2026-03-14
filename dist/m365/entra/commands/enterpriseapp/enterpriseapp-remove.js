var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraEnterpriseAppRemoveCommand_instances, _EntraEnterpriseAppRemoveCommand_initTelemetry, _EntraEnterpriseAppRemoveCommand_initOptions, _EntraEnterpriseAppRemoveCommand_initValidators, _EntraEnterpriseAppRemoveCommand_initOptionSets, _EntraEnterpriseAppRemoveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { odata } from '../../../../utils/odata.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraEnterpriseAppRemoveCommand extends GraphCommand {
    get name() {
        return commands.ENTERPRISEAPP_REMOVE;
    }
    get description() {
        return 'Deletes an enterprise application (or service principal)';
    }
    alias() {
        return [commands.SP_REMOVE];
    }
    constructor() {
        super();
        _EntraEnterpriseAppRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraEnterpriseAppRemoveCommand_instances, "m", _EntraEnterpriseAppRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraEnterpriseAppRemoveCommand_instances, "m", _EntraEnterpriseAppRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraEnterpriseAppRemoveCommand_instances, "m", _EntraEnterpriseAppRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraEnterpriseAppRemoveCommand_instances, "m", _EntraEnterpriseAppRemoveCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _EntraEnterpriseAppRemoveCommand_instances, "m", _EntraEnterpriseAppRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const removeEnterpriseApplication = async () => {
            if (this.verbose) {
                await logger.logToStderr(`Removing enterprise application ${args.options.id || args.options.displayName || args.options.objectId}...`);
            }
            try {
                let url = `${this.resource}/v1.0`;
                if (args.options.id) {
                    url += `/servicePrincipals(appId='${args.options.id}')`;
                }
                else {
                    const id = await this.getSpId(args.options);
                    url += `/servicePrincipals/${id}`;
                }
                const requestOptions = {
                    url: url,
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
            await removeEnterpriseApplication();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove enterprise application '${args.options.id || args.options.displayName || args.options.objectId}'?` });
            if (result) {
                await removeEnterpriseApplication();
            }
        }
    }
    async getSpId(options) {
        if (options.objectId) {
            return options.objectId;
        }
        const spItemsResponse = await odata.getAllItems(`${this.resource}/v1.0/servicePrincipals?$filter=displayName eq '${formatting.encodeQueryParameter(options.displayName)}'&$select=id`);
        if (spItemsResponse.length === 0) {
            throw `The specified enterprise application does not exist.`;
        }
        if (spItemsResponse.length > 1) {
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', spItemsResponse);
            const result = await cli.handleMultipleResultsFound(`Multiple enterprise applications with name '${options.displayName}' found.`, resultAsKeyValuePair);
            return result.id;
        }
        const spItem = spItemsResponse[0];
        return spItem.id;
    }
}
_EntraEnterpriseAppRemoveCommand_instances = new WeakSet(), _EntraEnterpriseAppRemoveCommand_initTelemetry = function _EntraEnterpriseAppRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            displayName: typeof args.options.displayName !== 'undefined',
            objectId: typeof args.options.objectId !== 'undefined',
            force: !!args.options.force
        });
    });
}, _EntraEnterpriseAppRemoveCommand_initOptions = function _EntraEnterpriseAppRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-n, --displayName [displayName]'
    }, {
        option: '--objectId [objectId]'
    }, {
        option: '-f, --force'
    });
}, _EntraEnterpriseAppRemoveCommand_initValidators = function _EntraEnterpriseAppRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `The option 'id' with value '${args.options.id}' is not a valid GUID.`;
        }
        if (args.options.objectId && !validation.isValidGuid(args.options.objectId)) {
            return `The option 'objectId' with value '${args.options.objectId}' is not a valid GUID.`;
        }
        return true;
    });
}, _EntraEnterpriseAppRemoveCommand_initOptionSets = function _EntraEnterpriseAppRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'displayName', 'objectId'] });
}, _EntraEnterpriseAppRemoveCommand_initTypes = function _EntraEnterpriseAppRemoveCommand_initTypes() {
    this.types.string.push('id', 'displayName', 'objectId');
    this.types.boolean.push('force');
};
export default new EntraEnterpriseAppRemoveCommand();
//# sourceMappingURL=enterpriseapp-remove.js.map