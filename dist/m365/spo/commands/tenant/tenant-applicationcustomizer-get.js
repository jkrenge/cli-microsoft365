var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoTenantApplicationCustomizerGetCommand_instances, _SpoTenantApplicationCustomizerGetCommand_initTelemetry, _SpoTenantApplicationCustomizerGetCommand_initOptions, _SpoTenantApplicationCustomizerGetCommand_initValidators, _SpoTenantApplicationCustomizerGetCommand_initOptionSets, _SpoTenantApplicationCustomizerGetCommand_initTypes;
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';
class SpoTenantApplicationCustomizerGetCommand extends SpoCommand {
    get name() {
        return commands.TENANT_APPLICATIONCUSTOMIZER_GET;
    }
    get description() {
        return 'Get an application customizer that is installed tenant wide';
    }
    constructor() {
        super();
        _SpoTenantApplicationCustomizerGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoTenantApplicationCustomizerGetCommand_instances, "m", _SpoTenantApplicationCustomizerGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoTenantApplicationCustomizerGetCommand_instances, "m", _SpoTenantApplicationCustomizerGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoTenantApplicationCustomizerGetCommand_instances, "m", _SpoTenantApplicationCustomizerGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoTenantApplicationCustomizerGetCommand_instances, "m", _SpoTenantApplicationCustomizerGetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoTenantApplicationCustomizerGetCommand_instances, "m", _SpoTenantApplicationCustomizerGetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            const appCatalogUrl = await spo.getTenantAppCatalogUrl(logger, this.debug);
            if (!appCatalogUrl) {
                throw 'No app catalog URL found';
            }
            let filter;
            if (args.options.title) {
                filter = `Title eq '${args.options.title}'`;
            }
            else if (args.options.id) {
                filter = `Id eq '${args.options.id}'`;
            }
            else {
                filter = `TenantWideExtensionComponentId eq '${args.options.clientSideComponentId}'`;
            }
            const listServerRelativeUrl = urlUtil.getServerRelativePath(appCatalogUrl, '/lists/TenantWideExtensions');
            const listItemInstances = await odata.getAllItems(`${appCatalogUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/items?$filter=TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer' and ${filter}`);
            if (listItemInstances) {
                if (listItemInstances.length === 0) {
                    throw 'The specified application customizer was not found';
                }
                listItemInstances.forEach(v => delete v['ID']);
                let listItemInstance;
                if (listItemInstances.length > 1) {
                    const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', listItemInstances);
                    listItemInstance = await cli.handleMultipleResultsFound(`Multiple application customizers with ${args.options.title || args.options.clientSideComponentId} were found.`, resultAsKeyValuePair);
                }
                else {
                    listItemInstance = listItemInstances[0];
                }
                if (!args.options.tenantWideExtensionComponentProperties) {
                    await logger.log(listItemInstance);
                }
                else {
                    const properties = formatting.tryParseJson(listItemInstance.TenantWideExtensionComponentProperties);
                    await logger.log(properties);
                }
            }
            else {
                throw 'The specified application customizer was not found';
            }
        }
        catch (err) {
            return this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoTenantApplicationCustomizerGetCommand_instances = new WeakSet(), _SpoTenantApplicationCustomizerGetCommand_initTelemetry = function _SpoTenantApplicationCustomizerGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            title: typeof args.options.title !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            clientSideComponentId: typeof args.options.clientSideComponentId !== 'undefined',
            tenantWideExtensionComponentProperties: !!args.options.tenantWideExtensionComponentProperties
        });
    });
}, _SpoTenantApplicationCustomizerGetCommand_initOptions = function _SpoTenantApplicationCustomizerGetCommand_initOptions() {
    this.options.unshift({
        option: '-t, --title [title]'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-c, --clientSideComponentId [clientSideComponentId]'
    }, {
        option: '-p, --tenantWideExtensionComponentProperties'
    });
}, _SpoTenantApplicationCustomizerGetCommand_initValidators = function _SpoTenantApplicationCustomizerGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && isNaN(parseInt(args.options.id))) {
            return `${args.options.id} is not a number`;
        }
        if (args.options.clientSideComponentId && !validation.isValidGuid(args.options.clientSideComponentId)) {
            return `${args.options.clientSideComponentId} is not a valid GUID`;
        }
        return true;
    });
}, _SpoTenantApplicationCustomizerGetCommand_initOptionSets = function _SpoTenantApplicationCustomizerGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['title', 'id', 'clientSideComponentId'] });
}, _SpoTenantApplicationCustomizerGetCommand_initTypes = function _SpoTenantApplicationCustomizerGetCommand_initTypes() {
    this.types.string.push('title', 'id', 'clientSideComponentId');
    this.types.boolean.push('tenantWideExtensionComponentProperties');
};
export default new SpoTenantApplicationCustomizerGetCommand();
//# sourceMappingURL=tenant-applicationcustomizer-get.js.map