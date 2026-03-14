var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoTenantCommandSetRemoveCommand_instances, _SpoTenantCommandSetRemoveCommand_initTelemetry, _SpoTenantCommandSetRemoveCommand_initOptions, _SpoTenantCommandSetRemoveCommand_initValidators, _SpoTenantCommandSetRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoTenantCommandSetRemoveCommand extends SpoCommand {
    get name() {
        return commands.TENANT_COMMANDSET_REMOVE;
    }
    get description() {
        return 'Removes a ListView Command Set that is installed tenant wide.';
    }
    constructor() {
        super();
        _SpoTenantCommandSetRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoTenantCommandSetRemoveCommand_instances, "m", _SpoTenantCommandSetRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoTenantCommandSetRemoveCommand_instances, "m", _SpoTenantCommandSetRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoTenantCommandSetRemoveCommand_instances, "m", _SpoTenantCommandSetRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoTenantCommandSetRemoveCommand_instances, "m", _SpoTenantCommandSetRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (args.options.force) {
                return await this.removeTenantCommandSet(logger, args);
            }
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the tenant commandset ${args.options.id || args.options.title || args.options.clientSideComponentId}?` });
            if (result) {
                await this.removeTenantCommandSet(logger, args);
            }
        }
        catch (err) {
            await logger.log(err);
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async removeTenantCommandSet(logger, args) {
        const appCatalogUrl = await spo.getTenantAppCatalogUrl(logger, this.debug);
        if (!appCatalogUrl) {
            throw 'No app catalog URL found';
        }
        const listServerRelativeUrl = urlUtil.getServerRelativePath(appCatalogUrl, '/lists/TenantWideExtensions');
        const requestUrl = `${appCatalogUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
        const id = await this.getTenantCommandSetId(logger, args, requestUrl);
        if (this.verbose) {
            await logger.logToStderr(`Removing tenant command set ${args.options.id || args.options.title || args.options.clientSideComponentId}`);
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
    async getTenantCommandSetId(logger, args, requestUrl) {
        if (this.verbose) {
            await logger.logToStderr(`Getting the tenant command set ${args.options.id || args.options.title || args.options.clientSideComponentId}`);
        }
        let filter;
        if (args.options.title) {
            filter = `Title eq '${args.options.title}'`;
        }
        else if (args.options.id) {
            filter = `Id eq ${args.options.id}`;
        }
        else {
            filter = `TenantWideExtensionComponentId eq '${args.options.clientSideComponentId}'`;
        }
        const listItemInstances = await odata.getAllItems(`${requestUrl}/items?$filter=startswith(TenantWideExtensionLocation,'ClientSideExtension.ListViewCommandSet') and ${filter}`);
        if (listItemInstances.length === 0) {
            throw 'The specified command set was not found';
        }
        if (listItemInstances.length > 1) {
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', listItemInstances);
            const result = await cli.handleMultipleResultsFound(`Multiple command sets with ${args.options.title || args.options.clientSideComponentId} were found.`, resultAsKeyValuePair);
            return result.Id;
        }
        return listItemInstances[0].Id;
    }
}
_SpoTenantCommandSetRemoveCommand_instances = new WeakSet(), _SpoTenantCommandSetRemoveCommand_initTelemetry = function _SpoTenantCommandSetRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            title: typeof args.options.title !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            clientSideComponentId: typeof args.options.clientSideComponentId !== 'undefined',
            force: !!args.options.force
        });
    });
}, _SpoTenantCommandSetRemoveCommand_initOptions = function _SpoTenantCommandSetRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-t, --title [title]'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-c, --clientSideComponentId  [clientSideComponentId]'
    }, {
        option: '-f, --force'
    });
}, _SpoTenantCommandSetRemoveCommand_initValidators = function _SpoTenantCommandSetRemoveCommand_initValidators() {
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
}, _SpoTenantCommandSetRemoveCommand_initOptionSets = function _SpoTenantCommandSetRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['title', 'id', 'clientSideComponentId'] });
};
export default new SpoTenantCommandSetRemoveCommand();
//# sourceMappingURL=tenant-commandset-remove.js.map