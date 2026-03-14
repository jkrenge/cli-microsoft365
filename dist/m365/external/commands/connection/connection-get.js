var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _ExternalConnectionGetCommand_instances, _ExternalConnectionGetCommand_initTelemetry, _ExternalConnectionGetCommand_initOptions, _ExternalConnectionGetCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class ExternalConnectionGetCommand extends GraphCommand {
    get name() {
        return commands.CONNECTION_GET;
    }
    get description() {
        return 'Get a specific external connection';
    }
    alias() {
        return [commands.EXTERNALCONNECTION_GET];
    }
    constructor() {
        super();
        _ExternalConnectionGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _ExternalConnectionGetCommand_instances, "m", _ExternalConnectionGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _ExternalConnectionGetCommand_instances, "m", _ExternalConnectionGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _ExternalConnectionGetCommand_instances, "m", _ExternalConnectionGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        let url = `${this.resource}/v1.0/external/connections`;
        if (args.options.id) {
            url += `/${formatting.encodeQueryParameter(args.options.id)}`;
        }
        else {
            url += `?$filter=name eq '${formatting.encodeQueryParameter(args.options.name)}'`;
        }
        const requestOptions = {
            url: url,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        try {
            let res = await request.get(requestOptions);
            if (args.options.name) {
                if (res.value.length === 0) {
                    throw `External connection with name '${args.options.name}' not found`;
                }
                res = res.value[0];
            }
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_ExternalConnectionGetCommand_instances = new WeakSet(), _ExternalConnectionGetCommand_initTelemetry = function _ExternalConnectionGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined'
        });
    });
}, _ExternalConnectionGetCommand_initOptions = function _ExternalConnectionGetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-n, --name [name]'
    });
}, _ExternalConnectionGetCommand_initOptionSets = function _ExternalConnectionGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] });
};
export default new ExternalConnectionGetCommand();
//# sourceMappingURL=connection-get.js.map