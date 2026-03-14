var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PpDataverseTableRowListCommand_instances, _PpDataverseTableRowListCommand_initTelemetry, _PpDataverseTableRowListCommand_initOptions, _PpDataverseTableRowListCommand_initOptionSets;
import request from '../../../../request.js';
import { odata } from '../../../../utils/odata.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
class PpDataverseTableRowListCommand extends PowerPlatformCommand {
    get name() {
        return commands.DATAVERSE_TABLE_ROW_LIST;
    }
    get description() {
        return 'Lists table rows for the given Dataverse table';
    }
    constructor() {
        super();
        _PpDataverseTableRowListCommand_instances.add(this);
        __classPrivateFieldGet(this, _PpDataverseTableRowListCommand_instances, "m", _PpDataverseTableRowListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PpDataverseTableRowListCommand_instances, "m", _PpDataverseTableRowListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PpDataverseTableRowListCommand_instances, "m", _PpDataverseTableRowListCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr('Retrieving list of table rows');
        }
        try {
            const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);
            const entitySetName = await this.getEntitySetName(dynamicsApiUrl, args);
            if (this.verbose) {
                await logger.logToStderr('Entity set name is: ' + entitySetName);
            }
            const response = await odata.getAllItems(`${dynamicsApiUrl}/api/data/v9.0/${entitySetName}`);
            await logger.log(response);
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
_PpDataverseTableRowListCommand_instances = new WeakSet(), _PpDataverseTableRowListCommand_initTelemetry = function _PpDataverseTableRowListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            entitySetName: typeof args.options.entitySetName !== 'undefined',
            tableName: typeof args.options.tableName !== 'undefined',
            asAdmin: !!args.options.asAdmin
        });
    });
}, _PpDataverseTableRowListCommand_initOptions = function _PpDataverseTableRowListCommand_initOptions() {
    this.options.unshift({
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '--entitySetName [entitySetName]'
    }, {
        option: '--tableName [tableName]'
    }, {
        option: '--asAdmin'
    });
}, _PpDataverseTableRowListCommand_initOptionSets = function _PpDataverseTableRowListCommand_initOptionSets() {
    this.optionSets.push({ options: ['entitySetName', 'tableName'] });
};
export default new PpDataverseTableRowListCommand();
//# sourceMappingURL=dataverse-table-row-list.js.map