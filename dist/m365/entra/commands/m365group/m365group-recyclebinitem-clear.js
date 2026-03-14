var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraM365GroupRecycleBinItemClearCommand_instances, _EntraM365GroupRecycleBinItemClearCommand_initTelemetry, _EntraM365GroupRecycleBinItemClearCommand_initOptions;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraM365GroupRecycleBinItemClearCommand extends GraphCommand {
    get name() {
        return commands.M365GROUP_RECYCLEBINITEM_CLEAR;
    }
    get description() {
        return 'Clears all M365 Groups from recycle bin.';
    }
    constructor() {
        super();
        _EntraM365GroupRecycleBinItemClearCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraM365GroupRecycleBinItemClearCommand_instances, "m", _EntraM365GroupRecycleBinItemClearCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupRecycleBinItemClearCommand_instances, "m", _EntraM365GroupRecycleBinItemClearCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        const clearM365GroupRecycleBinItems = async () => {
            try {
                await this.processRecycleBinItemsClear();
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await clearM365GroupRecycleBinItems();
        }
        else {
            const response = await cli.promptForConfirmation({ message: `Are you sure you want to clear all M365 Groups from recycle bin?` });
            if (response) {
                await clearM365GroupRecycleBinItems();
            }
        }
    }
    async processRecycleBinItemsClear() {
        const filter = `?$filter=groupTypes/any(c:c+eq+'Unified')`;
        const topCount = '&$top=100';
        const endpoint = `${this.resource}/v1.0/directory/deletedItems/Microsoft.Graph.Group${filter}${topCount}`;
        const recycleBinItems = await odata.getAllItems(endpoint);
        if (recycleBinItems.length === 0) {
            return;
        }
        const deletePromises = [];
        // Logic to delete a group from recycle bin items.
        recycleBinItems.forEach(grp => {
            deletePromises.push(request.delete({
                url: `${this.resource}/v1.0/directory/deletedItems/${grp.id}`,
                headers: {
                    'accept': 'application/json;odata.metadata=none'
                }
            }));
        });
        await Promise.all(deletePromises);
    }
}
_EntraM365GroupRecycleBinItemClearCommand_instances = new WeakSet(), _EntraM365GroupRecycleBinItemClearCommand_initTelemetry = function _EntraM365GroupRecycleBinItemClearCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: typeof args.options.force !== 'undefined'
        });
    });
}, _EntraM365GroupRecycleBinItemClearCommand_initOptions = function _EntraM365GroupRecycleBinItemClearCommand_initOptions() {
    this.options.unshift({
        option: '-f, --force'
    });
};
export default new EntraM365GroupRecycleBinItemClearCommand();
//# sourceMappingURL=m365group-recyclebinitem-clear.js.map