var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PpDataverseTableListCommand_instances, _PpDataverseTableListCommand_initTelemetry, _PpDataverseTableListCommand_initOptions;
import { odata } from '../../../../utils/odata.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
class PpDataverseTableListCommand extends PowerPlatformCommand {
    get name() {
        return commands.DATAVERSE_TABLE_LIST;
    }
    get description() {
        return 'Lists dataverse tables in a given environment';
    }
    defaultProperties() {
        return ['SchemaName', 'EntitySetName', 'LogicalName', 'IsManaged'];
    }
    constructor() {
        super();
        _PpDataverseTableListCommand_instances.add(this);
        __classPrivateFieldGet(this, _PpDataverseTableListCommand_instances, "m", _PpDataverseTableListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PpDataverseTableListCommand_instances, "m", _PpDataverseTableListCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving list of tables for which the user is an admin...`);
        }
        try {
            const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);
            const endpoint = `${dynamicsApiUrl}/api/data/v9.0/EntityDefinitions?$select=MetadataId,IsCustomEntity,IsManaged,SchemaName,IconVectorName,LogicalName,EntitySetName,IsActivity,DataProviderId,IsRenameable,IsCustomizable,CanCreateForms,CanCreateViews,CanCreateCharts,CanCreateAttributes,CanChangeTrackingBeEnabled,CanModifyAdditionalSettings,CanChangeHierarchicalRelationship,CanEnableSyncToExternalSearchIndex&$filter=(IsIntersect eq false and IsLogicalEntity eq false and%0APrimaryNameAttribute ne null and PrimaryNameAttribute ne %27%27 and ObjectTypeCode gt 0 and%0AObjectTypeCode ne 4712 and ObjectTypeCode ne 4724 and ObjectTypeCode ne 9933 and ObjectTypeCode ne 9934 and%0AObjectTypeCode ne 9935 and ObjectTypeCode ne 9947 and ObjectTypeCode ne 9945 and ObjectTypeCode ne 9944 and%0AObjectTypeCode ne 9942 and ObjectTypeCode ne 9951 and ObjectTypeCode ne 2016 and ObjectTypeCode ne 9949 and%0AObjectTypeCode ne 9866 and ObjectTypeCode ne 9867 and ObjectTypeCode ne 9868) and (IsCustomizable/Value eq true or IsCustomEntity eq true or IsManaged eq false or IsMappable/Value eq true or IsRenameable/Value eq true)&api-version=9.1`;
            const res = await odata.getAllItems(endpoint);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_PpDataverseTableListCommand_instances = new WeakSet(), _PpDataverseTableListCommand_initTelemetry = function _PpDataverseTableListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            asAdmin: !!args.options.asAdmin
        });
    });
}, _PpDataverseTableListCommand_initOptions = function _PpDataverseTableListCommand_initOptions() {
    this.options.unshift({
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '--asAdmin'
    });
};
export default new PpDataverseTableListCommand();
//# sourceMappingURL=dataverse-table-list.js.map