var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraEnterpriseAppAddCommand_instances, _EntraEnterpriseAppAddCommand_initTelemetry, _EntraEnterpriseAppAddCommand_initOptions, _EntraEnterpriseAppAddCommand_initValidators, _EntraEnterpriseAppAddCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';
class EntraEnterpriseAppAddCommand extends GraphCommand {
    get name() {
        return commands.ENTERPRISEAPP_ADD;
    }
    get description() {
        return 'Creates an enterprise application (or service principal) for a registered Entra app';
    }
    alias() {
        return [commands.SP_ADD];
    }
    constructor() {
        super();
        _EntraEnterpriseAppAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraEnterpriseAppAddCommand_instances, "m", _EntraEnterpriseAppAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraEnterpriseAppAddCommand_instances, "m", _EntraEnterpriseAppAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraEnterpriseAppAddCommand_instances, "m", _EntraEnterpriseAppAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraEnterpriseAppAddCommand_instances, "m", _EntraEnterpriseAppAddCommand_initOptionSets).call(this);
    }
    async getAppId(args) {
        if (args.options.id) {
            return args.options.id;
        }
        let spMatchQuery = '';
        if (args.options.displayName) {
            spMatchQuery = `displayName eq '${formatting.encodeQueryParameter(args.options.displayName)}'`;
        }
        else if (args.options.objectId) {
            spMatchQuery = `id eq '${formatting.encodeQueryParameter(args.options.objectId)}'`;
        }
        const appIdRequestOptions = {
            url: `${this.resource}/v1.0/applications?$filter=${spMatchQuery}`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const response = await request.get(appIdRequestOptions);
        const spItem = response.value[0];
        if (!spItem) {
            throw `The specified Entra app doesn't exist`;
        }
        if (response.value.length > 1) {
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('appId', response.value);
            const result = await cli.handleMultipleResultsFound(`Multiple Entra apps with name '${args.options.displayName}' found.`, resultAsKeyValuePair);
            return result.appId;
        }
        return spItem.appId;
    }
    async commandAction(logger, args) {
        try {
            const appId = await this.getAppId(args);
            const requestOptions = {
                url: `${this.resource}/v1.0/servicePrincipals`,
                headers: {
                    accept: 'application/json;odata.metadata=none',
                    'content-type': 'application/json;odata=nometadata'
                },
                data: {
                    appId: appId
                },
                responseType: 'json'
            };
            const res = await request.post(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_EntraEnterpriseAppAddCommand_instances = new WeakSet(), _EntraEnterpriseAppAddCommand_initTelemetry = function _EntraEnterpriseAppAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: (!(!args.options.id)).toString(),
            displayName: (!(!args.options.displayName)).toString(),
            objectId: (!(!args.options.objectId)).toString()
        });
    });
}, _EntraEnterpriseAppAddCommand_initOptions = function _EntraEnterpriseAppAddCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-n, --displayName [displayName]'
    }, {
        option: '--objectId [objectId]'
    });
}, _EntraEnterpriseAppAddCommand_initValidators = function _EntraEnterpriseAppAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        if (args.options.objectId && !validation.isValidGuid(args.options.objectId)) {
            return `${args.options.objectId} is not a valid GUID`;
        }
        return true;
    });
}, _EntraEnterpriseAppAddCommand_initOptionSets = function _EntraEnterpriseAppAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'displayName', 'objectId'] });
};
export default new EntraEnterpriseAppAddCommand();
//# sourceMappingURL=enterpriseapp-add.js.map