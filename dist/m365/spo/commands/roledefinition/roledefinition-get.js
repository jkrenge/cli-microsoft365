var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoRoleDefinitionGetCommand_instances, _SpoRoleDefinitionGetCommand_initOptions, _SpoRoleDefinitionGetCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import { BasePermissions } from '../../base-permissions.js';
import commands from '../../commands.js';
import { RoleType } from './RoleType.js';
class SpoRoleDefinitionGetCommand extends SpoCommand {
    get name() {
        return commands.ROLEDEFINITION_GET;
    }
    get description() {
        return 'Gets specified role definition from web by id';
    }
    constructor() {
        super();
        _SpoRoleDefinitionGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoRoleDefinitionGetCommand_instances, "m", _SpoRoleDefinitionGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoRoleDefinitionGetCommand_instances, "m", _SpoRoleDefinitionGetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Getting role definition from ${args.options.webUrl}...`);
        }
        const requestOptions = {
            url: `${args.options.webUrl}/_api/web/roledefinitions(${args.options.id})`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        try {
            const response = await request.get(requestOptions);
            const permissions = new BasePermissions();
            permissions.high = response.BasePermissions.High;
            permissions.low = response.BasePermissions.Low;
            response.BasePermissionsValue = permissions.parse();
            response.RoleTypeKindValue = RoleType[response.RoleTypeKind];
            await logger.log(response);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoRoleDefinitionGetCommand_instances = new WeakSet(), _SpoRoleDefinitionGetCommand_initOptions = function _SpoRoleDefinitionGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-i, --id <id>'
    });
}, _SpoRoleDefinitionGetCommand_initValidators = function _SpoRoleDefinitionGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (isNaN(args.options.id)) {
            return `${args.options.id} is not a number`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
};
export default new SpoRoleDefinitionGetCommand();
//# sourceMappingURL=roledefinition-get.js.map