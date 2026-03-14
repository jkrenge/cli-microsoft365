var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoRoleDefinitionListCommand_instances, _SpoRoleDefinitionListCommand_initOptions, _SpoRoleDefinitionListCommand_initValidators;
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoRoleDefinitionListCommand extends SpoCommand {
    get name() {
        return commands.ROLEDEFINITION_LIST;
    }
    get description() {
        return 'Gets list of role definitions for the specified site';
    }
    defaultProperties() {
        return ['Id', 'Name'];
    }
    constructor() {
        super();
        _SpoRoleDefinitionListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoRoleDefinitionListCommand_instances, "m", _SpoRoleDefinitionListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoRoleDefinitionListCommand_instances, "m", _SpoRoleDefinitionListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Getting role definitions list from ${args.options.webUrl}...`);
        }
        try {
            const res = await odata.getAllItems(`${args.options.webUrl}/_api/web/roledefinitions`);
            const response = formatting.setFriendlyPermissions(res);
            await logger.log(response);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoRoleDefinitionListCommand_instances = new WeakSet(), _SpoRoleDefinitionListCommand_initOptions = function _SpoRoleDefinitionListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    });
}, _SpoRoleDefinitionListCommand_initValidators = function _SpoRoleDefinitionListCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.webUrl));
};
export default new SpoRoleDefinitionListCommand();
//# sourceMappingURL=roledefinition-list.js.map