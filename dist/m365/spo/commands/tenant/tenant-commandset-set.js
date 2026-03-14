var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoTenantCommandSetSetCommand_instances, _a, _SpoTenantCommandSetSetCommand_initTelemetry, _SpoTenantCommandSetSetCommand_initOptions, _SpoTenantCommandSetSetCommand_initValidators;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoTenantCommandSetSetCommand extends SpoCommand {
    get name() {
        return commands.TENANT_COMMANDSET_SET;
    }
    get description() {
        return 'Updates a ListView Command Set that is installed tenant wide.';
    }
    constructor() {
        super();
        _SpoTenantCommandSetSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoTenantCommandSetSetCommand_instances, "m", _SpoTenantCommandSetSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoTenantCommandSetSetCommand_instances, "m", _SpoTenantCommandSetSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoTenantCommandSetSetCommand_instances, "m", _SpoTenantCommandSetSetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const appCatalogUrl = await spo.getTenantAppCatalogUrl(logger, this.debug);
            if (!appCatalogUrl) {
                throw 'No app catalog URL found';
            }
            const listServerRelativeUrl = urlUtil.getServerRelativePath(appCatalogUrl, '/lists/TenantWideExtensions');
            const listItem = await this.getListItemById(logger, appCatalogUrl, listServerRelativeUrl, args.options.id);
            if (listItem.TenantWideExtensionLocation.indexOf("ClientSideExtension.ListViewCommandSet") === -1) {
                throw 'The item is not a ListViewCommandSet';
            }
            await this.updateTenantWideExtension(appCatalogUrl, args.options, listServerRelativeUrl, logger);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getListItemById(logger, webUrl, listServerRelativeUrl, id) {
        if (this.verbose) {
            await logger.logToStderr(`Getting the list item by id ${id}`);
        }
        const reqOptions = {
            url: `${webUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/Items(${id})`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        return await request.get(reqOptions);
    }
    async updateTenantWideExtension(appCatalogUrl, options, listServerRelativeUrl, logger) {
        if (this.verbose) {
            await logger.logToStderr('Updating tenant wide extension to the TenantWideExtensions list');
        }
        const formValues = [];
        if (options.newTitle !== undefined) {
            formValues.push({
                FieldName: 'Title',
                FieldValue: options.newTitle
            });
        }
        if (options.clientSideComponentId !== undefined) {
            formValues.push({
                FieldName: 'TenantWideExtensionComponentId',
                FieldValue: options.clientSideComponentId
            });
        }
        if (options.location !== undefined) {
            formValues.push({
                FieldName: 'TenantWideExtensionLocation',
                FieldValue: this.getLocation(options.location)
            });
        }
        if (options.listType !== undefined) {
            formValues.push({
                FieldName: 'TenantWideExtensionListTemplate',
                FieldValue: this.getListTemplate(options.listType)
            });
        }
        if (options.clientSideComponentProperties !== undefined) {
            formValues.push({
                FieldName: 'TenantWideExtensionComponentProperties',
                FieldValue: options.clientSideComponentProperties
            });
        }
        if (options.webTemplate !== undefined) {
            formValues.push({
                FieldName: 'TenantWideExtensionWebTemplate',
                FieldValue: options.webTemplate
            });
        }
        const requestOptions = {
            url: `${appCatalogUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/Items(${options.id})/ValidateUpdateListItem()`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            data: {
                formValues: formValues
            },
            responseType: 'json'
        };
        await request.post(requestOptions);
    }
    getLocation(location) {
        switch (location) {
            case 'Both':
                return 'ClientSideExtension.ListViewCommandSet';
            case 'ContextMenu':
                return 'ClientSideExtension.ListViewCommandSet.ContextMenu';
            default:
                return 'ClientSideExtension.ListViewCommandSet.CommandBar';
        }
    }
    getListTemplate(listTemplate) {
        switch (listTemplate) {
            case 'SitePages':
                return '119';
            case 'Library':
                return '101';
            default:
                return '100';
        }
    }
}
_a = SpoTenantCommandSetSetCommand, _SpoTenantCommandSetSetCommand_instances = new WeakSet(), _SpoTenantCommandSetSetCommand_initTelemetry = function _SpoTenantCommandSetSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            newTitle: typeof args.options.newTitle !== 'undefined',
            listType: args.options.listType,
            clientSideComponentId: typeof args.options.clientSideComponentId !== 'undefined',
            clientSideComponentProperties: typeof args.options.clientSideComponentProperties !== 'undefined',
            webTemplate: typeof args.options.webTemplate !== 'undefined',
            location: args.options.location
        });
    });
}, _SpoTenantCommandSetSetCommand_initOptions = function _SpoTenantCommandSetSetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '-t, --newTitle [newTitle]'
    }, {
        option: '-l, --listType [listType]',
        autocomplete: _a.listTypes
    }, {
        option: '-c, --clientSideComponentId [clientSideComponentId]'
    }, {
        option: '-p, --clientSideComponentProperties [clientSideComponentProperties]'
    }, {
        option: '-w, --webTemplate [webTemplate]'
    }, {
        option: '--location [location]',
        autocomplete: _a.locations
    });
}, _SpoTenantCommandSetSetCommand_initValidators = function _SpoTenantCommandSetSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!args.options.newTitle &&
            !args.options.listType &&
            !args.options.clientSideComponentId &&
            !args.options.clientSideComponentProperties &&
            !args.options.webTemplate &&
            !args.options.location) {
            return 'Specify at least one property to update';
        }
        if (args.options.clientSideComponentId && !validation.isValidGuid(args.options.clientSideComponentId)) {
            return `${args.options.clientSideComponentId} is not a valid GUID`;
        }
        if (args.options.listType && _a.listTypes.indexOf(args.options.listType) < 0) {
            return `${args.options.listType} is not a valid list type. Allowed values are ${_a.listTypes.join(', ')}`;
        }
        if (args.options.location && _a.locations.indexOf(args.options.location) < 0) {
            return `${args.options.location} is not a valid location. Allowed values are ${_a.locations.join(', ')}`;
        }
        return true;
    });
};
SpoTenantCommandSetSetCommand.listTypes = ['List', 'Library', 'SitePages'];
SpoTenantCommandSetSetCommand.locations = ['ContextMenu', 'CommandBar', 'Both'];
export default new SpoTenantCommandSetSetCommand();
//# sourceMappingURL=tenant-commandset-set.js.map