var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraM365GroupRenewCommand_instances, _EntraM365GroupRenewCommand_initOptions, _EntraM365GroupRenewCommand_initValidators, _EntraM365GroupRenewCommand_initOptionSets, _EntraM365GroupRenewCommand_initTypes;
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraM365GroupRenewCommand extends GraphCommand {
    get name() {
        return commands.M365GROUP_RENEW;
    }
    get description() {
        return `Renews Microsoft 365 group's expiration`;
    }
    constructor() {
        super();
        _EntraM365GroupRenewCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraM365GroupRenewCommand_instances, "m", _EntraM365GroupRenewCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupRenewCommand_instances, "m", _EntraM365GroupRenewCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupRenewCommand_instances, "m", _EntraM365GroupRenewCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupRenewCommand_instances, "m", _EntraM365GroupRenewCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Renewing Microsoft 365 group's expiration: ${args.options.id || args.options.displayName}...`);
        }
        try {
            let groupId = args.options.id;
            if (args.options.displayName) {
                groupId = await entraGroup.getGroupIdByDisplayName(args.options.displayName);
            }
            const isUnifiedGroup = await entraGroup.isUnifiedGroup(groupId);
            if (!isUnifiedGroup) {
                throw Error(`Specified group with id '${groupId}' is not a Microsoft 365 group.`);
            }
            const requestOptions = {
                url: `${this.resource}/v1.0/groups/${groupId}/renew/`,
                headers: {
                    'accept': 'application/json;odata.metadata=none'
                }
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_EntraM365GroupRenewCommand_instances = new WeakSet(), _EntraM365GroupRenewCommand_initOptions = function _EntraM365GroupRenewCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-n, --displayName [displayName]'
    });
}, _EntraM365GroupRenewCommand_initValidators = function _EntraM365GroupRenewCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
}, _EntraM365GroupRenewCommand_initOptionSets = function _EntraM365GroupRenewCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'displayName'] });
}, _EntraM365GroupRenewCommand_initTypes = function _EntraM365GroupRenewCommand_initTypes() {
    this.types.string.push('id', 'displayName');
};
export default new EntraM365GroupRenewCommand();
//# sourceMappingURL=m365group-renew.js.map