var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TenantIdGetCommand_instances, _TenantIdGetCommand_initTelemetry, _TenantIdGetCommand_initOptions;
import auth from '../../../../Auth.js';
import Command from '../../../../Command.js';
import request from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import commands from '../../commands.js';
class TenantIdGetCommand extends Command {
    get name() {
        return commands.ID_GET;
    }
    get description() {
        return 'Gets Microsoft 365 tenant ID for the specified domain';
    }
    constructor() {
        super();
        _TenantIdGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _TenantIdGetCommand_instances, "m", _TenantIdGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TenantIdGetCommand_instances, "m", _TenantIdGetCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        let domainName = args.options.domainName;
        if (!domainName) {
            const userName = accessToken.getUserNameFromAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken);
            domainName = userName.split('@')[1];
        }
        const requestOptions = {
            url: `https://login.windows.net/${domainName}/.well-known/openid-configuration`,
            headers: {
                'content-type': 'application/json',
                accept: 'application/json',
                'x-anonymous': true
            },
            responseType: 'json'
        };
        try {
            const res = await request.get(requestOptions);
            if (res.error) {
                throw res.error_description;
            }
            if (res.token_endpoint) {
                await logger.log(res.token_endpoint.split('/')[3]);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_TenantIdGetCommand_instances = new WeakSet(), _TenantIdGetCommand_initTelemetry = function _TenantIdGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            domainName: typeof args.options.domainName !== 'undefined'
        });
    });
}, _TenantIdGetCommand_initOptions = function _TenantIdGetCommand_initOptions() {
    this.options.unshift({
        option: '-d, --domainName [domainName]'
    });
};
export default new TenantIdGetCommand();
//# sourceMappingURL=id-get.js.map