var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListItemBatchAddCommand_instances, _SpoListItemBatchAddCommand_initTelemetry, _SpoListItemBatchAddCommand_initOptions, _SpoListItemBatchAddCommand_initValidators, _SpoListItemBatchAddCommand_initTypes, _SpoListItemBatchAddCommand_initOptionSets;
import fs from 'fs';
import os from 'os';
import { v4 } from 'uuid';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListItemBatchAddCommand extends SpoCommand {
    allowUnknownOptions() {
        return true;
    }
    get name() {
        return commands.LISTITEM_BATCH_ADD;
    }
    get description() {
        return 'Creates list items in a batch';
    }
    constructor() {
        super();
        _SpoListItemBatchAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListItemBatchAddCommand_instances, "m", _SpoListItemBatchAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListItemBatchAddCommand_instances, "m", _SpoListItemBatchAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListItemBatchAddCommand_instances, "m", _SpoListItemBatchAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListItemBatchAddCommand_instances, "m", _SpoListItemBatchAddCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _SpoListItemBatchAddCommand_instances, "m", _SpoListItemBatchAddCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Starting to create batch items from csv ${args.options.filePath ? `at path ${args.options.filePath}` : 'from content'}`);
            }
            const csvContent = args.options.filePath ? fs.readFileSync(args.options.filePath, 'utf8') : args.options.csvContent;
            const jsonContent = formatting.parseCsvToJson(csvContent);
            await this.addItemsAsBatch(jsonContent, args.options, logger);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async addItemsAsBatch(rows, options, logger) {
        const requestUrl = this.getRequestUrl(options);
        let itemsToAdd = [];
        for await (const [index, row] of rows.entries()) {
            itemsToAdd.push(this.getSingleItemRequestBody(row));
            if (itemsToAdd.length === 100) {
                if (this.verbose) {
                    await logger.logToStderr(`Writing away batch of items, currently at: ${index + 1}/${rows.length}.`);
                }
                await this.postBatchData(itemsToAdd, options.webUrl, requestUrl);
                itemsToAdd = [];
            }
        }
        if (itemsToAdd.length) {
            if (this.verbose) {
                await logger.logToStderr(`Writing away ${itemsToAdd.length} items.`);
            }
            await this.postBatchData(itemsToAdd, options.webUrl, requestUrl);
        }
    }
    async postBatchData(itemsToAdd, webUrl, requestUrl) {
        const batchId = v4();
        const requestBody = this.parseBatchRequestBody(itemsToAdd, batchId, requestUrl);
        const requestOptions = {
            url: `${webUrl}/_api/$batch`,
            headers: {
                'Content-Type': `multipart/mixed; boundary=batch_${batchId}`,
                'Accept': 'application/json;odata=verbose'
            },
            data: requestBody.join('')
        };
        const response = await request.post(requestOptions);
        const parsedResponse = this.parseBatchResponseBody(response);
        if (parsedResponse.some(r => r.HasException)) {
            throw `Creating some items failed with the following errors: ${os.EOL}${parsedResponse.filter(f => f.HasException).map(f => { return `- Line ${f.csvLineNumber}: ${f.FieldName} - ${f.ErrorMessage}`; }).join(os.EOL)}`;
        }
    }
    parseBatchRequestBody(items, batchId, requestUrl) {
        const changeSetId = v4();
        const batchBody = [];
        // add default batch body headers
        batchBody.push(`--batch_${batchId}\n`);
        batchBody.push(`Content-Type: multipart/mixed; boundary="changeset_${changeSetId}"\n\n`);
        batchBody.push('Content-Transfer-Encoding: binary\n\n');
        items.forEach((item) => {
            batchBody.push(`--changeset_${changeSetId}\n`);
            batchBody.push('Content-Type: application/http\n');
            batchBody.push('Content-Transfer-Encoding: binary\n\n');
            batchBody.push(`POST ${requestUrl} HTTP/1.1\n`);
            batchBody.push(`Accept: application/json;odata=nometadata\n`);
            batchBody.push(`Content-Type: application/json;odata=verbose\n`);
            batchBody.push(`If-Match: *\n\n`);
            batchBody.push(`{\n"formValues": ${JSON.stringify(this.formatFormValues(item))}\n}`);
        });
        // close batch body
        batchBody.push(`\n\n`);
        batchBody.push(`--changeset_${changeSetId}--\n\n`);
        batchBody.push(`--batch_${batchId}--\n`);
        return batchBody;
    }
    formatFormValues(input) {
        // Fix for PS 7
        return input.map(obj => ({
            FieldName: obj.FieldName.replace(/\\"/g, ''),
            FieldValue: obj.FieldValue.replace(/\\"/g, '')
        }));
    }
    parseBatchResponseBody(response) {
        const batchResults = [];
        response.split('\r\n')
            .filter((line) => line.startsWith('{'))
            .forEach((line, index) => {
            const parsedResponse = JSON.parse(line);
            if (parsedResponse.error) {
                // if an error object is returned, the request failed
                const error = parsedResponse.error;
                throw error.message.value;
            }
            parsedResponse.value.forEach((fieldValueResult) => {
                batchResults.push({
                    csvLineNumber: (index + 2),
                    ...fieldValueResult
                });
            });
        });
        return batchResults;
    }
    getSingleItemRequestBody(row) {
        const requestBody = [];
        Object.keys(row).forEach(key => {
            // have to do 'toString()' or the API will complain when entering a numeric field
            requestBody.push({ FieldName: key, FieldValue: row[key].toString() });
        });
        return requestBody;
    }
    getRequestUrl(options) {
        let listUrl = `${options.webUrl}/_api/web`;
        if (options.listId) {
            listUrl += `/lists(guid'${formatting.encodeQueryParameter(options.listId)}')`;
        }
        else if (options.listTitle) {
            listUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(options.listTitle)}')`;
        }
        else if (options.listUrl) {
            const listServerRelativeUrl = urlUtil.getServerRelativePath(options.webUrl, options.listUrl);
            listUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
        }
        return `${listUrl}/AddValidateUpdateItemUsingPath`;
    }
}
_SpoListItemBatchAddCommand_instances = new WeakSet(), _SpoListItemBatchAddCommand_initTelemetry = function _SpoListItemBatchAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            filePath: typeof args.options.filePath !== 'undefined',
            csvContent: typeof args.options.csvContent !== 'undefined',
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined'
        });
    });
}, _SpoListItemBatchAddCommand_initOptions = function _SpoListItemBatchAddCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-p, --filePath [filePath]'
    }, {
        option: '-c, --csvContent [csvContent]'
    }, {
        option: '-l, --listId [listId]'
    }, {
        option: '-t, --listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    });
}, _SpoListItemBatchAddCommand_initValidators = function _SpoListItemBatchAddCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.listId &&
            !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} in option listId is not a valid GUID`;
        }
        if (args.options.filePath && !fs.existsSync(args.options.filePath)) {
            return `File with path ${args.options.filePath} does not exist`;
        }
        return true;
    });
}, _SpoListItemBatchAddCommand_initTypes = function _SpoListItemBatchAddCommand_initTypes() {
    this.types.string.push('webUrl', 'filePath', 'listId', 'listTitle', 'listUrl');
}, _SpoListItemBatchAddCommand_initOptionSets = function _SpoListItemBatchAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] }, { options: ['filePath', 'csvContent'] });
};
export default new SpoListItemBatchAddCommand();
//# sourceMappingURL=listitem-batch-add.js.map