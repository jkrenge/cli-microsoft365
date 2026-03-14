var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListItemRetentionLabelEnsureCommand_instances, _SpoListItemRetentionLabelEnsureCommand_initTelemetry, _SpoListItemRetentionLabelEnsureCommand_initOptions, _SpoListItemRetentionLabelEnsureCommand_initValidators, _SpoListItemRetentionLabelEnsureCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { spo } from '../../../../utils/spo.js';
import { odata } from '../../../../utils/odata.js';
class SpoListItemRetentionLabelEnsureCommand extends SpoCommand {
    get name() {
        return commands.LISTITEM_RETENTIONLABEL_ENSURE;
    }
    get description() {
        return 'Apply a retention label to a list item';
    }
    constructor() {
        super();
        _SpoListItemRetentionLabelEnsureCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListItemRetentionLabelEnsureCommand_instances, "m", _SpoListItemRetentionLabelEnsureCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListItemRetentionLabelEnsureCommand_instances, "m", _SpoListItemRetentionLabelEnsureCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListItemRetentionLabelEnsureCommand_instances, "m", _SpoListItemRetentionLabelEnsureCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListItemRetentionLabelEnsureCommand_instances, "m", _SpoListItemRetentionLabelEnsureCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const listAbsoluteUrl = await this.getListAbsoluteUrl(args.options, logger);
            const labelName = await this.getLabelName(args.options, logger);
            if (args.options.assetId) {
                await this.applyAssetId(args.options, logger);
            }
            await spo.applyRetentionLabelToListItems(args.options.webUrl, labelName, listAbsoluteUrl, [parseInt(args.options.listItemId)], logger, args.options.verbose);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getLabelName(options, logger) {
        if (options.name) {
            return options.name;
        }
        if (this.verbose) {
            await logger.logToStderr(`Retrieving the name of the retention label based on the Id '${options.id}'...`);
        }
        const requestUrl = `${options.webUrl}/_api/SP.CompliancePolicy.SPPolicyStoreProxy.GetAvailableTagsForSite(siteUrl=@a1)?@a1='${formatting.encodeQueryParameter(options.webUrl)}'`;
        const labels = await odata.getAllItems(requestUrl);
        const label = labels.find(l => l.TagId === options.id);
        if (label === undefined) {
            throw new Error(`The specified retention label does not exist or is not published to this SharePoint site. Use the name of the label if you want to apply an unpublished label.`);
        }
        if (this.verbose && label !== undefined) {
            await logger.logToStderr(`Retention label found in the list of available labels: '${label.TagName}' / '${label.TagId}'...`);
        }
        return label.TagName;
    }
    async getListAbsoluteUrl(options, logger) {
        const parsedUrl = new URL(options.webUrl);
        const tenantUrl = `${parsedUrl.protocol}//${parsedUrl.hostname}`;
        if (options.listUrl) {
            const serverRelativePath = urlUtil.getServerRelativePath(options.webUrl, options.listUrl);
            return urlUtil.urlCombine(tenantUrl, serverRelativePath);
        }
        if (this.verbose) {
            await logger.logToStderr(`Retrieving list absolute URL...`);
        }
        let requestUrl = `${options.webUrl}/_api/web`;
        if (options.listId) {
            requestUrl += `/lists(guid'${formatting.encodeQueryParameter(options.listId)}')`;
        }
        else if (options.listTitle) {
            requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(options.listTitle)}')`;
        }
        const requestOptions = {
            url: `${requestUrl}?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        const serverRelativePath = urlUtil.getServerRelativePath(options.webUrl, response.RootFolder.ServerRelativeUrl);
        const listAbsoluteUrl = urlUtil.urlCombine(tenantUrl, serverRelativePath);
        if (this.verbose) {
            await logger.logToStderr(`List absolute URL found: '${listAbsoluteUrl}'`);
        }
        return listAbsoluteUrl;
    }
    async applyAssetId(options, logger) {
        if (this.verbose) {
            await logger.logToStderr(`Applying the asset Id ${options.assetId}...`);
        }
        let requestUrl = `${options.webUrl}/_api/web`;
        if (options.listId) {
            requestUrl += `/lists(guid'${formatting.encodeQueryParameter(options.listId)}')/items(${options.listItemId})/ValidateUpdateListItem()`;
        }
        else if (options.listTitle) {
            requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(options.listTitle)}')/items(${options.listItemId})/ValidateUpdateListItem()`;
        }
        else if (options.listUrl) {
            const listServerRelativeUrl = urlUtil.getServerRelativePath(options.webUrl, options.listUrl);
            requestUrl += `/GetList(@listUrl)/items(${options.listItemId})/ValidateUpdateListItem()?@listUrl='${formatting.encodeQueryParameter(listServerRelativeUrl)}'`;
        }
        const requestBody = { "formValues": [{ "FieldName": "ComplianceAssetId", "FieldValue": options.assetId }] };
        const requestOptions = {
            url: requestUrl,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            data: requestBody,
            responseType: 'json'
        };
        await request.post(requestOptions);
    }
}
_SpoListItemRetentionLabelEnsureCommand_instances = new WeakSet(), _SpoListItemRetentionLabelEnsureCommand_initTelemetry = function _SpoListItemRetentionLabelEnsureCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            assetId: typeof args.options.assetId !== 'undefined'
        });
    });
}, _SpoListItemRetentionLabelEnsureCommand_initOptions = function _SpoListItemRetentionLabelEnsureCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--listItemId <listItemId>'
    }, {
        option: '--listId [listId]'
    }, {
        option: '--listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '-n, --name [name]'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-a, --assetId [assetId]'
    });
}, _SpoListItemRetentionLabelEnsureCommand_initValidators = function _SpoListItemRetentionLabelEnsureCommand_initValidators() {
    this.validators.push(async (args) => {
        const id = parseInt(args.options.listItemId);
        if (isNaN(id)) {
            return `${args.options.listItemId} is not a valid list item ID`;
        }
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.listId &&
            !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} is not a valid GUID`;
        }
        if (args.options.id &&
            !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
}, _SpoListItemRetentionLabelEnsureCommand_initOptionSets = function _SpoListItemRetentionLabelEnsureCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
    this.optionSets.push({ options: ['name', 'id'] });
};
export default new SpoListItemRetentionLabelEnsureCommand();
//# sourceMappingURL=listitem-retentionlabel-ensure.js.map