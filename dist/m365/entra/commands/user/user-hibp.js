var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraUserHibpCommand_instances, _EntraUserHibpCommand_initTelemetry, _EntraUserHibpCommand_initOptions, _EntraUserHibpCommand_initValidators;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import AnonymousCommand from '../../../base/AnonymousCommand.js';
import commands from '../../commands.js';
class EntraUserHibpCommand extends AnonymousCommand {
    get name() {
        return commands.USER_HIBP;
    }
    get description() {
        return 'Allows you to retrieve all accounts that have been pwned with the specified username';
    }
    constructor() {
        super();
        _EntraUserHibpCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraUserHibpCommand_instances, "m", _EntraUserHibpCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraUserHibpCommand_instances, "m", _EntraUserHibpCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraUserHibpCommand_instances, "m", _EntraUserHibpCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const requestOptions = {
                url: `https://haveibeenpwned.com/api/v3/breachedaccount/${formatting.encodeQueryParameter(args.options.userName)}${(args.options.domain ? `?domain=${formatting.encodeQueryParameter(args.options.domain)}` : '')}`,
                headers: {
                    'accept': 'application/json',
                    'hibp-api-key': args.options.apiKey,
                    'x-anonymous': true
                },
                responseType: 'json'
            };
            const res = await request.get(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            if ((err && err.response !== undefined && err.response.status === 404) && (this.debug || this.verbose)) {
                await logger.log('No pwnage found');
                return;
            }
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_EntraUserHibpCommand_instances = new WeakSet(), _EntraUserHibpCommand_initTelemetry = function _EntraUserHibpCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            domain: args.options.domain
        });
    });
}, _EntraUserHibpCommand_initOptions = function _EntraUserHibpCommand_initOptions() {
    this.options.unshift({
        option: '-n, --userName <userName>'
    }, {
        option: '--apiKey, <apiKey>'
    }, {
        option: '--domain, [domain]'
    });
}, _EntraUserHibpCommand_initValidators = function _EntraUserHibpCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidUserPrincipalName(args.options.userName)) {
            return 'Specify valid userName';
        }
        return true;
    });
};
export default new EntraUserHibpCommand();
//# sourceMappingURL=user-hibp.js.map