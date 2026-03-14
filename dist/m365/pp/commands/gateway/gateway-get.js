var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PpGatewayGetCommand_instances, _PpGatewayGetCommand_initOptions, _PpGatewayGetCommand_initValidators;
import request from "../../../../request.js";
import { formatting } from "../../../../utils/formatting.js";
import { validation } from "../../../../utils/validation.js";
import PowerBICommand from "../../../base/PowerBICommand.js";
import commands from "../../commands.js";
class PpGatewayGetCommand extends PowerBICommand {
    get name() {
        return commands.GATEWAY_GET;
    }
    get description() {
        return "Get information about the specified gateway.";
    }
    constructor() {
        super();
        _PpGatewayGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _PpGatewayGetCommand_instances, "m", _PpGatewayGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PpGatewayGetCommand_instances, "m", _PpGatewayGetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const gateway = await this.getGateway(args.options.id);
            await logger.log(gateway);
        }
        catch (error) {
            this.handleRejectedODataJsonPromise(error);
        }
    }
    getGateway(gatewayId) {
        const requestOptions = {
            url: `${this.resource}/v1.0/myorg/gateways/${formatting.encodeQueryParameter(gatewayId)}`,
            headers: {
                accept: "application/json;odata.metadata=none"
            },
            responseType: "json"
        };
        return request.get(requestOptions);
    }
}
_PpGatewayGetCommand_instances = new WeakSet(), _PpGatewayGetCommand_initOptions = function _PpGatewayGetCommand_initOptions() {
    this.options.unshift({
        option: "-i, --id <id>"
    });
}, _PpGatewayGetCommand_initValidators = function _PpGatewayGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
};
export default new PpGatewayGetCommand();
//# sourceMappingURL=gateway-get.js.map