var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoTenantApplicationCustomizerRemoveCommand_instances, _SpoTenantApplicationCustomizerRemoveCommand_initTelemetry, _SpoTenantApplicationCustomizerRemoveCommand_initOptions, _SpoTenantApplicationCustomizerRemoveCommand_initValidators, _SpoTenantApplicationCustomizerRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoTenantApplicationCustomizerRemoveCommand extends SpoCommand {
    get name() {
        return commands.TENANT_APPLICATIONCUSTOMIZER_REMOVE;
    }
    get description() {
        return 'Removes an application customizer that is installed tenant wide.';
    }
    constructor() {
        super();
        _SpoTenantApplicationCustomizerRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoTenantApplicationCustomizerRemoveCommand_instances, "m", _SpoTenantApplicationCustomizerRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoTenantApplicationCustomizerRemoveCommand_instances, "m", _SpoTenantApplicationCustomizerRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoTenantApplicationCustomizerRemoveCommand_instances, "m", _SpoTenantApplicationCustomizerRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoTenantApplicationCustomizerRemoveCommand_instances, "m", _SpoTenantApplicationCustomizerRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (args.options.force) {
                return await this.removeTenantApplicationCustomizer(logger, args);
            }
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the tenant applicationcustomizer ${args.options.id || args.options.title || args.options.clientSideComponentId}?` });
            if (result) {
                await this.removeTenantApplicationCustomizer(logger, args);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getTenantApplicationCustomizerId(logger, args, requestUrl) {
        if (this.verbose) {
            await logger.logToStderr(`Getting the tenant application customizer ${args.options.id || args.options.title || args.options.clientSideComponentId}`);
        }
        const filter = [`TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer'`];
        if (args.options.title) {
            filter.push(`Title eq '${args.options.title}'`);
        }
        else if (args.options.id) {
            filter.push(`Id eq '${args.options.id}'`);
        }
        else if (args.options.clientSideComponentId) {
            filter.push(`TenantWideExtensionComponentId eq '${args.options.clientSideComponentId}'`);
        }
        const listItemInstances = await odata.getAllItems(`${requestUrl}/items?$filter=${filter.join(' and ')}&$select=Id`);
        if (listItemInstances.length === 0) {
            throw 'The specified application customizer was not found';
        }
        if (listItemInstances.length > 1) {
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', listItemInstances);
            listItemInstances[0] = await cli.handleMultipleResultsFound(`Multiple application customizers with ${args.options.title || args.options.clientSideComponentId} were found.`, resultAsKeyValuePair);
        }
        return listItemInstances[0].Id;
    }
    async removeTenantApplicationCustomizer(logger, args) {
        const appCatalogUrl = await spo.getTenantAppCatalogUrl(logger, this.debug);
        if (!appCatalogUrl) {
            throw 'No app catalog URL found';
        }
        const listServerRelativeUrl = urlUtil.getServerRelativePath(appCatalogUrl, '/lists/TenantWideExtensions');
        const requestUrl = `${appCatalogUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
        const id = await this.getTenantApplicationCustomizerId(logger, args, requestUrl);
        if (this.verbose) {
            await logger.logToStderr(`Removing tenant application customizer ${args.options.id || args.options.title || args.options.clientSideComponentId}`);
        }
        const requestOptions = {
            url: `${requestUrl}/items(${id})`,
            method: 'POST',
            headers: {
                'X-HTTP-Method': 'DELETE',
                'If-Match': '*',
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        await request.post(requestOptions);
    }
}
_SpoTenantApplicationCustomizerRemoveCommand_instances = new WeakSet(), _SpoTenantApplicationCustomizerRemoveCommand_initTelemetry = function _SpoTenantApplicationCustomizerRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            title: typeof args.options.title !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            clientSideComponentId: typeof args.options.clientSideComponentId !== 'undefined',
            force: !!args.options.force
        });
    });
}, _SpoTenantApplicationCustomizerRemoveCommand_initOptions = function _SpoTenantApplicationCustomizerRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-t, --title [title]'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-c, --clientSideComponentId  [clientSideComponentId]'
    }, {
        option: '-f, --force'
    });
}, _SpoTenantApplicationCustomizerRemoveCommand_initValidators = function _SpoTenantApplicationCustomizerRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id) {
            const id = parseInt(args.options.id);
            if (isNaN(id)) {
                return `${args.options.id} is not a valid list item ID`;
            }
        }
        if (args.options.clientSideComponentId &&
            !validation.isValidGuid(args.options.clientSideComponentId)) {
            return `${args.options.clientSideComponentId} is not a valid GUID`;
        }
        return true;
    });
}, _SpoTenantApplicationCustomizerRemoveCommand_initOptionSets = function _SpoTenantApplicationCustomizerRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['title', 'id', 'clientSideComponentId'] });
};
export default new SpoTenantApplicationCustomizerRemoveCommand();
//# sourceMappingURL=tenant-applicationcustomizer-remove.js.map