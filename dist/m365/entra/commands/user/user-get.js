var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraUserGetCommand_instances, _EntraUserGetCommand_initTelemetry, _EntraUserGetCommand_initOptions, _EntraUserGetCommand_initValidators, _EntraUserGetCommand_initOptionSets;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { formatting } from '../../../../utils/formatting.js';
class EntraUserGetCommand extends GraphCommand {
    get name() {
        return commands.USER_GET;
    }
    get description() {
        return 'Gets information about the specified user';
    }
    constructor() {
        super();
        _EntraUserGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraUserGetCommand_instances, "m", _EntraUserGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraUserGetCommand_instances, "m", _EntraUserGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraUserGetCommand_instances, "m", _EntraUserGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraUserGetCommand_instances, "m", _EntraUserGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            let userIdOrPrincipalName = args.options.id;
            if (args.options.userName) {
                // single user can be retrieved also by user principal name
                userIdOrPrincipalName = formatting.encodeQueryParameter(args.options.userName);
            }
            else if (args.options.email) {
                userIdOrPrincipalName = await entraUser.getUserIdByEmail(args.options.email);
            }
            const requestUrl = this.getRequestUrl(userIdOrPrincipalName, args.options);
            const requestOptions = {
                url: requestUrl,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            const user = await request.get(requestOptions);
            await logger.log(user);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getRequestUrl(userIdOrPrincipalName, options) {
        const queryParameters = [];
        if (options.properties) {
            const allProperties = options.properties.split(',');
            const selectProperties = allProperties.filter(prop => !prop.includes('/'));
            if (selectProperties.length > 0) {
                queryParameters.push(`$select=${selectProperties}`);
            }
        }
        if (options.withManager) {
            queryParameters.push('$expand=manager($select=displayName,userPrincipalName,id,mail)');
        }
        const queryString = queryParameters.length > 0
            ? `?${queryParameters.join('&')}`
            : '';
        // user principal name can start with $ but it violates the OData URL convention, so it must be enclosed in parenthesis and single quotes
        return userIdOrPrincipalName.startsWith('%24')
            ? `${this.resource}/v1.0/users('${userIdOrPrincipalName}')${queryString}`
            : `${this.resource}/v1.0/users/${userIdOrPrincipalName}${queryString}`;
    }
}
_EntraUserGetCommand_instances = new WeakSet(), _EntraUserGetCommand_initTelemetry = function _EntraUserGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            email: typeof args.options.email !== 'undefined',
            properties: args.options.properties,
            withManager: typeof args.options.withManager !== 'undefined'
        });
    });
}, _EntraUserGetCommand_initOptions = function _EntraUserGetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-n, --userName [userName]'
    }, {
        option: '--email [email]'
    }, {
        option: '-p, --properties [properties]'
    }, {
        option: '--withManager'
    });
}, _EntraUserGetCommand_initValidators = function _EntraUserGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id &&
            !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid userName`;
        }
        return true;
    });
}, _EntraUserGetCommand_initOptionSets = function _EntraUserGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'userName', 'email'] });
};
export default new EntraUserGetCommand();
//# sourceMappingURL=user-get.js.map