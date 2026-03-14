var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _VivaEngageUserGetCommand_instances, _VivaEngageUserGetCommand_initTelemetry, _VivaEngageUserGetCommand_initOptions, _VivaEngageUserGetCommand_initValidators;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import VivaEngageCommand from '../../../base/VivaEngageCommand.js';
import commands from '../../commands.js';
class VivaEngageUserGetCommand extends VivaEngageCommand {
    get name() {
        return commands.ENGAGE_USER_GET;
    }
    get description() {
        return 'Retrieves the current user or searches for a user by ID or e-mail';
    }
    constructor() {
        super();
        _VivaEngageUserGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _VivaEngageUserGetCommand_instances, "m", _VivaEngageUserGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _VivaEngageUserGetCommand_instances, "m", _VivaEngageUserGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _VivaEngageUserGetCommand_instances, "m", _VivaEngageUserGetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        let endPoint = `${this.resource}/v1/users/current.json`;
        if (args.options.id) {
            endPoint = `${this.resource}/v1/users/${args.options.id}.json`;
        }
        else if (args.options.email) {
            endPoint = `${this.resource}/v1/users/by_email.json?email=${formatting.encodeQueryParameter(args.options.email)}`;
        }
        const requestOptions = {
            url: endPoint,
            headers: {
                accept: 'application/json;odata.metadata=none',
                'content-type': 'application/json;odata=nometadata'
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
_VivaEngageUserGetCommand_instances = new WeakSet(), _VivaEngageUserGetCommand_initTelemetry = function _VivaEngageUserGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: args.options.id !== undefined,
            email: args.options.email !== undefined
        });
    });
}, _VivaEngageUserGetCommand_initOptions = function _VivaEngageUserGetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '--email [email]'
    });
}, _VivaEngageUserGetCommand_initValidators = function _VivaEngageUserGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id !== undefined && args.options.email !== undefined) {
            return `You are only allowed to search by ID or e-mail but not both`;
        }
        return true;
    });
};
export default new VivaEngageUserGetCommand();
//# sourceMappingURL=engage-user-get.js.map