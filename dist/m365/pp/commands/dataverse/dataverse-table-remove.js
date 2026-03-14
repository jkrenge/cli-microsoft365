var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PpDataverseTableRemoveCommand_instances, _PpDataverseTableRemoveCommand_initTelemetry, _PpDataverseTableRemoveCommand_initOptions;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
class PpDataverseTableRemoveCommand extends PowerPlatformCommand {
    get name() {
        return commands.DATAVERSE_TABLE_REMOVE;
    }
    get description() {
        return 'Removes a dataverse table in a given environment';
    }
    constructor() {
        super();
        _PpDataverseTableRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _PpDataverseTableRemoveCommand_instances, "m", _PpDataverseTableRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PpDataverseTableRemoveCommand_instances, "m", _PpDataverseTableRemoveCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Removing a table for which the user is an admin...`);
        }
        if (args.options.force) {
            await this.removeDataverseTable(args.options);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the dataverse table ${args.options.name}?` });
            if (result) {
                await this.removeDataverseTable(args.options);
            }
        }
    }
    async removeDataverseTable(options) {
        try {
            const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(options.environmentName, options.asAdmin);
            const requestOptions = {
                url: `${dynamicsApiUrl}/api/data/v9.0/EntityDefinitions(LogicalName='${options.name}')`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            await request.delete(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_PpDataverseTableRemoveCommand_instances = new WeakSet(), _PpDataverseTableRemoveCommand_initTelemetry = function _PpDataverseTableRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            asAdmin: !!args.options.asAdmin,
            force: !!args.options.force
        });
    });
}, _PpDataverseTableRemoveCommand_initOptions = function _PpDataverseTableRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '-n, --name <name>'
    }, {
        option: '--asAdmin'
    }, {
        option: '-f, --force'
    });
};
export default new PpDataverseTableRemoveCommand();
//# sourceMappingURL=dataverse-table-remove.js.map