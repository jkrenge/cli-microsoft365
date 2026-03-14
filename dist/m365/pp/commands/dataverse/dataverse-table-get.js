var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PpDataverseTableGetCommand_instances, _PpDataverseTableGetCommand_initTelemetry, _PpDataverseTableGetCommand_initOptions;
import request from '../../../../request.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
class PpDataverseTableGetCommand extends PowerPlatformCommand {
    get name() {
        return commands.DATAVERSE_TABLE_GET;
    }
    get description() {
        return 'Gets a dataverse table in a given environment';
    }
    constructor() {
        super();
        _PpDataverseTableGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _PpDataverseTableGetCommand_instances, "m", _PpDataverseTableGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PpDataverseTableGetCommand_instances, "m", _PpDataverseTableGetCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving a table for which the user is an admin...`);
        }
        try {
            const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);
            const requestOptions = {
                url: `${dynamicsApiUrl}/api/data/v9.0/EntityDefinitions(LogicalName='${args.options.name}')?$select=MetadataId,IsCustomEntity,IsManaged,SchemaName,IconVectorName,LogicalName,EntitySetName,IsActivity,DataProviderId,IsRenameable,IsCustomizable,CanCreateForms,CanCreateViews,CanCreateCharts,CanCreateAttributes,CanChangeTrackingBeEnabled,CanModifyAdditionalSettings,CanChangeHierarchicalRelationship,CanEnableSyncToExternalSearchIndex&api-version=9.1`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            const res = await request.get(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_PpDataverseTableGetCommand_instances = new WeakSet(), _PpDataverseTableGetCommand_initTelemetry = function _PpDataverseTableGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            asAdmin: !!args.options.asAdmin
        });
    });
}, _PpDataverseTableGetCommand_initOptions = function _PpDataverseTableGetCommand_initOptions() {
    this.options.unshift({
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '-n, --name <name>'
    }, {
        option: '--asAdmin'
    });
};
export default new PpDataverseTableGetCommand();
//# sourceMappingURL=dataverse-table-get.js.map