var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TenantInfoGetCommand_instances, _TenantInfoGetCommand_initTelemetry, _TenantInfoGetCommand_initOptions, _TenantInfoGetCommand_initValidators;
import auth from '../../../../Auth.js';
import GraphCommand from '../../../base/GraphCommand.js';
import request from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
class TenantInfoGetCommand extends GraphCommand {
    get name() {
        return commands.INFO_GET;
    }
    get description() {
        return 'Gets information about any tenant';
    }
    constructor() {
        super();
        _TenantInfoGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _TenantInfoGetCommand_instances, "m", _TenantInfoGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TenantInfoGetCommand_instances, "m", _TenantInfoGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TenantInfoGetCommand_instances, "m", _TenantInfoGetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        let domainName = args.options.domainName;
        const tenantId = args.options.tenantId;
        if (!domainName && !tenantId) {
            const userName = accessToken.getUserNameFromAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken);
            domainName = userName.split('@')[1];
        }
        let requestUrl = `${this.resource}/v1.0/tenantRelationships/`;
        if (tenantId) {
            requestUrl += `findTenantInformationByTenantId(tenantId='${formatting.encodeQueryParameter(tenantId)}')`;
        }
        else {
            requestUrl += `findTenantInformationByDomainName(domainName='${formatting.encodeQueryParameter(domainName)}')`;
        }
        const requestOptions = {
            url: requestUrl,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        try {
            const res = await request.get(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_TenantInfoGetCommand_instances = new WeakSet(), _TenantInfoGetCommand_initTelemetry = function _TenantInfoGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            domainName: typeof args.options.domainName !== 'undefined',
            tenantId: typeof args.options.tenantId !== 'undefined'
        });
    });
}, _TenantInfoGetCommand_initOptions = function _TenantInfoGetCommand_initOptions() {
    this.options.unshift({
        option: '-d, --domainName [domainName]'
    }, {
        option: '-i, --tenantId [tenantId]'
    });
}, _TenantInfoGetCommand_initValidators = function _TenantInfoGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.tenantId && !validation.isValidGuid(args.options.tenantId)) {
            return `${args.options.tenantId} is not a valid GUID`;
        }
        if (args.options.tenantId && args.options.domainName) {
            return `Specify either domainName or tenantId but not both`;
        }
        return true;
    });
};
export default new TenantInfoGetCommand();
//# sourceMappingURL=info-get.js.map