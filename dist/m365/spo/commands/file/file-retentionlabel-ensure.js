var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFileRetentionLabelEnsureCommand_instances, _SpoFileRetentionLabelEnsureCommand_initTelemetry, _SpoFileRetentionLabelEnsureCommand_initOptions, _SpoFileRetentionLabelEnsureCommand_initValidators, _SpoFileRetentionLabelEnsureCommand_initOptionSets, _SpoFileRetentionLabelEnsureCommand_initTypes;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { spo } from '../../../../utils/spo.js';
class SpoFileRetentionLabelEnsureCommand extends SpoCommand {
    get name() {
        return commands.FILE_RETENTIONLABEL_ENSURE;
    }
    get description() {
        return 'Apply a retention label to a file';
    }
    constructor() {
        super();
        _SpoFileRetentionLabelEnsureCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFileRetentionLabelEnsureCommand_instances, "m", _SpoFileRetentionLabelEnsureCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFileRetentionLabelEnsureCommand_instances, "m", _SpoFileRetentionLabelEnsureCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFileRetentionLabelEnsureCommand_instances, "m", _SpoFileRetentionLabelEnsureCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFileRetentionLabelEnsureCommand_instances, "m", _SpoFileRetentionLabelEnsureCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFileRetentionLabelEnsureCommand_instances, "m", _SpoFileRetentionLabelEnsureCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            const fileProperties = await this.getFileProperties(logger, args);
            const parsedUrl = new URL(args.options.webUrl);
            const tenantUrl = `${parsedUrl.protocol}//${parsedUrl.hostname}`;
            const listAbsoluteUrl = urlUtil.urlCombine(tenantUrl, fileProperties.listServerRelativeUrl);
            if (args.options.assetId) {
                await this.applyAssetId(args.options.webUrl, fileProperties.listServerRelativeUrl, fileProperties.listItemId, args.options.assetId);
            }
            await spo.applyRetentionLabelToListItems(args.options.webUrl, args.options.name, listAbsoluteUrl, [parseInt(fileProperties.listItemId)], logger, args.options.verbose);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getFileProperties(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving list and item information for file '${args.options.fileId || args.options.fileUrl}' in site at ${args.options.webUrl}...`);
        }
        let requestUrl = `${args.options.webUrl}/_api/web/`;
        if (args.options.fileId) {
            requestUrl += `GetFileById('${args.options.fileId}')`;
        }
        else {
            const serverRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl);
            requestUrl += `GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}')`;
        }
        const requestOptions = {
            url: `${requestUrl}?$expand=ListItemAllFields,ListItemAllFields/ParentList/RootFolder&$select=ServerRelativeUrl,ListItemAllFields/ParentList/RootFolder/ServerRelativeUrl,ListItemAllFields/Id`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        return { listItemId: response.ListItemAllFields.Id, listServerRelativeUrl: response.ListItemAllFields.ParentList.RootFolder.ServerRelativeUrl };
    }
    async applyAssetId(webUrl, listServerRelativeUrl, listItemId, assetId) {
        const requestBody = { "formValues": [{ "FieldName": "ComplianceAssetId", "FieldValue": assetId }] };
        const requestOptions = {
            url: `${webUrl}/_api/web/GetList(@listUrl)/items(${listItemId})/ValidateUpdateListItem()?@listUrl='${formatting.encodeQueryParameter(listServerRelativeUrl)}'`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            data: requestBody,
            responseType: 'json'
        };
        return request.post(requestOptions);
    }
}
_SpoFileRetentionLabelEnsureCommand_instances = new WeakSet(), _SpoFileRetentionLabelEnsureCommand_initTelemetry = function _SpoFileRetentionLabelEnsureCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            fileUrl: typeof args.options.fileUrl !== 'undefined',
            fileId: typeof args.options.fileId !== 'undefined',
            assetId: typeof args.options.assetId !== 'undefined'
        });
    });
}, _SpoFileRetentionLabelEnsureCommand_initOptions = function _SpoFileRetentionLabelEnsureCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--name <name>'
    }, {
        option: '--fileUrl [fileUrl]'
    }, {
        option: '-i, --fileId [fileId]'
    }, {
        option: '-a, --assetId [assetId]'
    });
}, _SpoFileRetentionLabelEnsureCommand_initValidators = function _SpoFileRetentionLabelEnsureCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.fileId &&
            !validation.isValidGuid(args.options.fileId)) {
            return `${args.options.fileId} is not a valid GUID`;
        }
        return true;
    });
}, _SpoFileRetentionLabelEnsureCommand_initOptionSets = function _SpoFileRetentionLabelEnsureCommand_initOptionSets() {
    this.optionSets.push({ options: ['fileUrl', 'fileId'] });
}, _SpoFileRetentionLabelEnsureCommand_initTypes = function _SpoFileRetentionLabelEnsureCommand_initTypes() {
    this.types.string.push('webUrl', 'name', 'fileUrl', 'fileId', 'assetId');
};
export default new SpoFileRetentionLabelEnsureCommand();
//# sourceMappingURL=file-retentionlabel-ensure.js.map