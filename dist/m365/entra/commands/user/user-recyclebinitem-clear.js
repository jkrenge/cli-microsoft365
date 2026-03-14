var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraUserRecycleBinItemClearCommand_instances, _EntraUserRecycleBinItemClearCommand_initTelemetry, _EntraUserRecycleBinItemClearCommand_initOptions;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraUserRecycleBinItemClearCommand extends GraphCommand {
    get name() {
        return commands.USER_RECYCLEBINITEM_CLEAR;
    }
    get description() {
        return 'Removes all users from the tenant recycle bin';
    }
    constructor() {
        super();
        _EntraUserRecycleBinItemClearCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraUserRecycleBinItemClearCommand_instances, "m", _EntraUserRecycleBinItemClearCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraUserRecycleBinItemClearCommand_instances, "m", _EntraUserRecycleBinItemClearCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        const clearRecycleBinUsers = async () => {
            try {
                const users = await odata.getAllItems(`${this.resource}/v1.0/directory/deletedItems/microsoft.graph.user?$select=id`);
                if (this.verbose) {
                    await logger.logToStderr(`Amount of users to permanently delete: ${users.length}`);
                }
                const batchRequests = users.map((user, index) => {
                    return {
                        id: index,
                        method: 'DELETE',
                        url: `/directory/deletedItems/${user.id}`
                    };
                });
                for (let i = 0; i < batchRequests.length; i += 20) {
                    const batchRequestChunk = batchRequests.slice(i, i + 20);
                    if (this.verbose) {
                        await logger.logToStderr(`Deleting users: ${i + batchRequestChunk.length}/${users.length}`);
                    }
                    const requestOptions = {
                        url: `${this.resource}/v1.0/$batch`,
                        headers: {
                            accept: 'application/json',
                            'content-type': 'application/json'
                        },
                        responseType: 'json',
                        data: {
                            requests: batchRequestChunk
                        }
                    };
                    await request.post(requestOptions);
                }
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await clearRecycleBinUsers();
        }
        else {
            const result = await cli.promptForConfirmation({ message: 'Are you sure you want to permanently delete all deleted users?' });
            if (result) {
                await clearRecycleBinUsers();
            }
        }
    }
}
_EntraUserRecycleBinItemClearCommand_instances = new WeakSet(), _EntraUserRecycleBinItemClearCommand_initTelemetry = function _EntraUserRecycleBinItemClearCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: !!args.options.force
        });
    });
}, _EntraUserRecycleBinItemClearCommand_initOptions = function _EntraUserRecycleBinItemClearCommand_initOptions() {
    this.options.unshift({
        option: '-f, --force'
    });
};
export default new EntraUserRecycleBinItemClearCommand();
//# sourceMappingURL=user-recyclebinitem-clear.js.map