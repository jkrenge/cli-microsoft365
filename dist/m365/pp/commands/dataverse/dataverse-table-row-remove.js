var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PpDataverseTableRowRemoveCommand_instances, _PpDataverseTableRowRemoveCommand_initTelemetry, _PpDataverseTableRowRemoveCommand_initOptions, _PpDataverseTableRowRemoveCommand_initValidators, _PpDataverseTableRowRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
class PpDataverseTableRowRemoveCommand extends PowerPlatformCommand {
    get name() {
        return commands.DATAVERSE_TABLE_ROW_REMOVE;
    }
    get description() {
        return 'Removes a specific row from a dataverse table in the specified Power Platform environment.';
    }
    constructor() {
        super();
        _PpDataverseTableRowRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _PpDataverseTableRowRemoveCommand_instances, "m", _PpDataverseTableRowRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PpDataverseTableRowRemoveCommand_instances, "m", _PpDataverseTableRowRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PpDataverseTableRowRemoveCommand_instances, "m", _PpDataverseTableRowRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PpDataverseTableRowRemoveCommand_instances, "m", _PpDataverseTableRowRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Removing row '${args.options.id}' from table '${args.options.tableName || args.options.entitySetName}'...`);
        }
        if (args.options.force) {
            await this.deleteTableRow(logger, args);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove row '${args.options.id}' from table '${args.options.tableName || args.options.entitySetName}'?` });
            if (result) {
                await this.deleteTableRow(logger, args);
            }
        }
    }
    async deleteTableRow(logger, args) {
        try {
            const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);
            const entitySetName = await this.getEntitySetName(dynamicsApiUrl, args);
            if (this.verbose) {
                await logger.logToStderr('Entity set name is: ' + entitySetName);
            }
            const requestOptions = {
                url: `${dynamicsApiUrl}/api/data/v9.1/${entitySetName}(${args.options.id})`,
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
    async getEntitySetName(dynamicsApiUrl, args) {
        if (args.options.entitySetName) {
            return args.options.entitySetName;
        }
        const requestOptions = {
            url: `${dynamicsApiUrl}/api/data/v9.0/EntityDefinitions(LogicalName='${args.options.tableName}')?$select=EntitySetName`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        return response.EntitySetName;
    }
}
_PpDataverseTableRowRemoveCommand_instances = new WeakSet(), _PpDataverseTableRowRemoveCommand_initTelemetry = function _PpDataverseTableRowRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            entitySetName: typeof args.options.entitySetName !== 'undefined',
            tableName: typeof args.options.tableName !== 'undefined',
            asAdmin: !!args.options.asAdmin,
            force: !!args.options.force
        });
    });
}, _PpDataverseTableRowRemoveCommand_initOptions = function _PpDataverseTableRowRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '-i, --id <id>'
    }, {
        option: '--entitySetName [entitySetName]'
    }, {
        option: '--tableName [tableName]'
    }, {
        option: '--asAdmin'
    }, {
        option: '-f, --force'
    });
}, _PpDataverseTableRowRemoveCommand_initValidators = function _PpDataverseTableRowRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
}, _PpDataverseTableRowRemoveCommand_initOptionSets = function _PpDataverseTableRowRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['entitySetName', 'tableName'] });
};
export default new PpDataverseTableRowRemoveCommand();
//# sourceMappingURL=dataverse-table-row-remove.js.map