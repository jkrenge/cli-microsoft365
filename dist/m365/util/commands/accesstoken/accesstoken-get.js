var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _UtilAccessTokenGetCommand_instances, _UtilAccessTokenGetCommand_initTelemetry, _UtilAccessTokenGetCommand_initOptions;
import auth, { Auth } from '../../../../Auth.js';
import Command from '../../../../Command.js';
import commands from '../../commands.js';
import { accessToken } from '../../../../utils/accessToken.js';
class UtilAccessTokenGetCommand extends Command {
    get name() {
        return commands.ACCESSTOKEN_GET;
    }
    get description() {
        return 'Gets access token for the specified resource';
    }
    constructor() {
        super();
        _UtilAccessTokenGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _UtilAccessTokenGetCommand_instances, "m", _UtilAccessTokenGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _UtilAccessTokenGetCommand_instances, "m", _UtilAccessTokenGetCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        let resource = args.options.resource;
        if (resource.toLowerCase() === 'sharepoint') {
            if (auth.connection.spoUrl) {
                resource = auth.connection.spoUrl;
            }
            else {
                throw `SharePoint URL undefined. Use the 'm365 spo set --url https://contoso.sharepoint.com' command to set the URL`;
            }
        }
        else if (resource.toLowerCase() === 'graph') {
            resource = Auth.getEndpointForResource('https://graph.microsoft.com', auth.connection.cloudType);
        }
        try {
            const token = await auth.ensureAccessToken(resource, logger, this.debug, args.options.new);
            if (args.options.decoded) {
                const { header, payload } = accessToken.getDecodedAccessToken(token);
                await logger.logRaw(`${JSON.stringify(header, null, 2)}.${JSON.stringify(payload, null, 2)}.[signature]`);
            }
            else {
                await logger.log(token);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_UtilAccessTokenGetCommand_instances = new WeakSet(), _UtilAccessTokenGetCommand_initTelemetry = function _UtilAccessTokenGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            new: args.options.new,
            decoded: args.options.decoded
        });
    });
}, _UtilAccessTokenGetCommand_initOptions = function _UtilAccessTokenGetCommand_initOptions() {
    this.options.unshift({
        option: '-r, --resource <resource>'
    }, {
        option: '--new'
    }, {
        option: '--decoded'
    });
};
export default new UtilAccessTokenGetCommand();
//# sourceMappingURL=accesstoken-get.js.map