var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraM365GroupRecycleBinItemListCommand_instances, _EntraM365GroupRecycleBinItemListCommand_initTelemetry, _EntraM365GroupRecycleBinItemListCommand_initOptions;
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraM365GroupRecycleBinItemListCommand extends GraphCommand {
    get name() {
        return commands.M365GROUP_RECYCLEBINITEM_LIST;
    }
    get description() {
        return 'Lists Microsoft 365 Groups deleted in the current tenant';
    }
    constructor() {
        super();
        _EntraM365GroupRecycleBinItemListCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraM365GroupRecycleBinItemListCommand_instances, "m", _EntraM365GroupRecycleBinItemListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupRecycleBinItemListCommand_instances, "m", _EntraM365GroupRecycleBinItemListCommand_initOptions).call(this);
    }
    defaultProperties() {
        return ['id', 'displayName', 'mailNickname'];
    }
    async commandAction(logger, args) {
        try {
            const filter = `?$filter=groupTypes/any(c:c+eq+'Unified')`;
            const displayNameFilter = args.options.groupName ? ` and startswith(DisplayName,'${formatting.encodeQueryParameter(args.options.groupName).replace(/'/g, `''`)}')` : '';
            const mailNicknameFilter = args.options.groupMailNickname ? ` and startswith(MailNickname,'${formatting.encodeQueryParameter(args.options.groupMailNickname).replace(/'/g, `''`)}')` : '';
            const topCount = '&$top=100';
            const endpoint = `${this.resource}/v1.0/directory/deletedItems/Microsoft.Graph.Group${filter}${displayNameFilter}${mailNicknameFilter}${topCount}`;
            const recycleBinItems = await odata.getAllItems(endpoint);
            await logger.log(recycleBinItems);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_EntraM365GroupRecycleBinItemListCommand_instances = new WeakSet(), _EntraM365GroupRecycleBinItemListCommand_initTelemetry = function _EntraM365GroupRecycleBinItemListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            groupName: typeof args.options.groupName !== 'undefined',
            groupMailNickname: typeof args.options.groupMailNickname !== 'undefined'
        });
    });
}, _EntraM365GroupRecycleBinItemListCommand_initOptions = function _EntraM365GroupRecycleBinItemListCommand_initOptions() {
    this.options.unshift({
        option: '-d, --groupName [groupName]'
    }, {
        option: '-m, --groupMailNickname [groupMailNickname]'
    });
};
export default new EntraM365GroupRecycleBinItemListCommand();
//# sourceMappingURL=m365group-recyclebinitem-list.js.map