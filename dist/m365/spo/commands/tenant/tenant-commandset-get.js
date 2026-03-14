var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoTenantCommandSetGetCommand_instances, _SpoTenantCommandSetGetCommand_initTelemetry, _SpoTenantCommandSetGetCommand_initOptions, _SpoTenantCommandSetGetCommand_initValidators, _SpoTenantCommandSetGetCommand_initOptionSets, _SpoTenantCommandSetGetCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoTenantCommandSetGetCommand extends SpoCommand {
    get name() {
        return commands.TENANT_COMMANDSET_GET;
    }
    get description() {
        return 'Get a ListView Command Set that is installed tenant wide';
    }
    constructor() {
        super();
        _SpoTenantCommandSetGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoTenantCommandSetGetCommand_instances, "m", _SpoTenantCommandSetGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoTenantCommandSetGetCommand_instances, "m", _SpoTenantCommandSetGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoTenantCommandSetGetCommand_instances, "m", _SpoTenantCommandSetGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoTenantCommandSetGetCommand_instances, "m", _SpoTenantCommandSetGetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoTenantCommandSetGetCommand_instances, "m", _SpoTenantCommandSetGetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const appCatalogUrl = await spo.getTenantAppCatalogUrl(logger, this.debug);
        if (!appCatalogUrl) {
            throw new CommandError('No app catalog URL found');
        }
        let filter = `startswith(TenantWideExtensionLocation,'ClientSideExtension.ListViewCommandSet')`;
        if (args.options.title) {
            filter += ` and Title eq '${args.options.title}'`;
        }
        else if (args.options.id) {
            filter += ` and Id eq ${args.options.id}`;
        }
        else if (args.options.clientSideComponentId) {
            filter += ` and TenantWideExtensionComponentId eq '${args.options.clientSideComponentId}'`;
        }
        const listServerRelativeUrl = urlUtil.getServerRelativePath(appCatalogUrl, '/lists/TenantWideExtensions');
        const reqOptions = {
            url: `${appCatalogUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/items?$filter=${filter}`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        try {
            const listItemInstances = await request.get(reqOptions);
            if (listItemInstances?.value.length > 0) {
                listItemInstances.value.forEach(v => delete v['ID']);
                let listItemInstance;
                if (listItemInstances.value.length > 1) {
                    const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', listItemInstances.value);
                    listItemInstance = await cli.handleMultipleResultsFound(`Multiple ListView Command Sets with ${args.options.title || args.options.clientSideComponentId} were found.`, resultAsKeyValuePair);
                }
                else {
                    listItemInstance = listItemInstances.value[0];
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
                throw 'The specified ListView Command Set was not found';
            }
        }
        catch (err) {
            return this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoTenantCommandSetGetCommand_instances = new WeakSet(), _SpoTenantCommandSetGetCommand_initTelemetry = function _SpoTenantCommandSetGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            title: typeof args.options.title !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            clientSideComponentId: typeof args.options.clientSideComponentId !== 'undefined',
            tenantWideExtensionComponentProperties: !!args.options.tenantWideExtensionComponentProperties
        });
    });
}, _SpoTenantCommandSetGetCommand_initOptions = function _SpoTenantCommandSetGetCommand_initOptions() {
    this.options.unshift({
        option: '-t, --title [title]'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-c, --clientSideComponentId  [clientSideComponentId]'
    }, {
        option: '-p, --tenantWideExtensionComponentProperties'
    });
}, _SpoTenantCommandSetGetCommand_initValidators = function _SpoTenantCommandSetGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && isNaN(parseInt(args.options.id))) {
            return `${args.options.id} is not a number`;
        }
        if (args.options.clientSideComponentId && !validation.isValidGuid(args.options.clientSideComponentId)) {
            return `${args.options.clientSideComponentId} is not a valid GUID`;
        }
        return true;
    });
}, _SpoTenantCommandSetGetCommand_initOptionSets = function _SpoTenantCommandSetGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['title', 'id', 'clientSideComponentId'] });
}, _SpoTenantCommandSetGetCommand_initTypes = function _SpoTenantCommandSetGetCommand_initTypes() {
    this.types.string.push('title', 'id', 'clientSideComponentId');
    this.types.boolean.push('tenantWideExtensionComponentProperties');
};
export default new SpoTenantCommandSetGetCommand();
//# sourceMappingURL=tenant-commandset-get.js.map