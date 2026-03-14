var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraEnterpriseAppGetCommand_instances, _EntraEnterpriseAppGetCommand_initTelemetry, _EntraEnterpriseAppGetCommand_initOptions, _EntraEnterpriseAppGetCommand_initValidators, _EntraEnterpriseAppGetCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraEnterpriseAppGetCommand extends GraphCommand {
    get name() {
        return commands.ENTERPRISEAPP_GET;
    }
    get description() {
        return 'Gets information about an Enterprise Application';
    }
    alias() {
        return [commands.SP_GET];
    }
    constructor() {
        super();
        _EntraEnterpriseAppGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraEnterpriseAppGetCommand_instances, "m", _EntraEnterpriseAppGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraEnterpriseAppGetCommand_instances, "m", _EntraEnterpriseAppGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraEnterpriseAppGetCommand_instances, "m", _EntraEnterpriseAppGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraEnterpriseAppGetCommand_instances, "m", _EntraEnterpriseAppGetCommand_initOptionSets).call(this);
    }
    async getSpId(args) {
        if (args.options.objectId) {
            return args.options.objectId;
        }
        let spMatchQuery = '';
        if (args.options.displayName) {
            spMatchQuery = `displayName eq '${formatting.encodeQueryParameter(args.options.displayName)}'`;
        }
        else if (args.options.id) {
            spMatchQuery = `appId eq '${formatting.encodeQueryParameter(args.options.id)}'`;
        }
        const idRequestOptions = {
            url: `${this.resource}/v1.0/servicePrincipals?$filter=${spMatchQuery}`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const response = await request.get(idRequestOptions);
        const spItem = response.value[0];
        if (!spItem) {
            throw `The specified Entra app does not exist`;
        }
        if (response.value.length > 1) {
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', response.value);
            const result = await cli.handleMultipleResultsFound(`Multiple Entra apps with name '${args.options.displayName}' found.`, resultAsKeyValuePair);
            return result.id;
        }
        return spItem.id;
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving enterprise application information...`);
        }
        try {
            const id = await this.getSpId(args);
            const requestOptions = {
                url: `${this.resource}/v1.0/servicePrincipals/${id}`,
                headers: {
                    accept: 'application/json;odata.metadata=none',
                    'content-type': 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            const res = await request.get(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_EntraEnterpriseAppGetCommand_instances = new WeakSet(), _EntraEnterpriseAppGetCommand_initTelemetry = function _EntraEnterpriseAppGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: (!(!args.options.id)).toString(),
            displayName: (!(!args.options.displayName)).toString(),
            objectId: (!(!args.options.objectId)).toString()
        });
    });
}, _EntraEnterpriseAppGetCommand_initOptions = function _EntraEnterpriseAppGetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-n, --displayName [displayName]'
    }, {
        option: '--objectId [objectId]'
    });
}, _EntraEnterpriseAppGetCommand_initValidators = function _EntraEnterpriseAppGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        if (args.options.objectId && !validation.isValidGuid(args.options.objectId)) {
            return `${args.options.objectId} is not a valid GUID`;
        }
        return true;
    });
}, _EntraEnterpriseAppGetCommand_initOptionSets = function _EntraEnterpriseAppGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'displayName', 'objectId'] });
};
export default new EntraEnterpriseAppGetCommand();
//# sourceMappingURL=enterpriseapp-get.js.map