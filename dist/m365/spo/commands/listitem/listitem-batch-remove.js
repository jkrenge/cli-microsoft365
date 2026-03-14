var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListItemBatchRemoveCommand_instances, _SpoListItemBatchRemoveCommand_initTelemetry, _SpoListItemBatchRemoveCommand_initOptions, _SpoListItemBatchRemoveCommand_initValidators, _SpoListItemBatchRemoveCommand_initTypes, _SpoListItemBatchRemoveCommand_initOptionSets;
import fs from 'fs';
import os from 'os';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { v4 } from 'uuid';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import { cli } from '../../../../cli/cli.js';
class SpoListItemBatchRemoveCommand extends SpoCommand {
    get name() {
        return commands.LISTITEM_BATCH_REMOVE;
    }
    get description() {
        return 'Removes items from a list in batch';
    }
    constructor() {
        super();
        _SpoListItemBatchRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListItemBatchRemoveCommand_instances, "m", _SpoListItemBatchRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListItemBatchRemoveCommand_instances, "m", _SpoListItemBatchRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListItemBatchRemoveCommand_instances, "m", _SpoListItemBatchRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListItemBatchRemoveCommand_instances, "m", _SpoListItemBatchRemoveCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _SpoListItemBatchRemoveCommand_instances, "m", _SpoListItemBatchRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const removeListItems = async () => {
            try {
                if (this.verbose) {
                    await logger.logToStderr('Removing the listitems from SharePoint...');
                }
                let idsToRemove = [];
                if (args.options.filePath) {
                    const csvContent = fs.readFileSync(args.options.filePath, 'utf-8');
                    const jsonContent = formatting.parseCsvToJson(csvContent);
                    const idKey = Object.keys(jsonContent[0]).find(key => key.toLowerCase() === 'id');
                    idsToRemove = jsonContent.map((item) => item[idKey]);
                }
                else {
                    idsToRemove = formatting.splitAndTrim(args.options.ids);
                }
                await this.removeItemsAsBatch(idsToRemove, args.options, logger);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeListItems();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to ${args.options.recycle ? "recycle" : "remove"} the list items from list ${args.options.listId || args.options.listTitle || args.options.listUrl} located in site ${args.options.webUrl}?` });
            if (result) {
                await removeListItems();
            }
        }
    }
    async removeItemsAsBatch(items, options, logger) {
        const itemsChunks = this.getChunkedArray(items, 10);
        for (const [index, chunk] of itemsChunks.entries()) {
            if (this.verbose) {
                await logger.logToStderr(`Processing chunk ${index + 1} of ${itemsChunks.length}...`);
            }
            await this.postBatchData(chunk, options.webUrl, options);
        }
    }
    async postBatchData(chunk, webUrl, options) {
        const batchId = v4();
        const requestBody = this.getRequestBody(chunk, batchId, options);
        const requestOptions = {
            url: `${webUrl}/_api/$batch`,
            headers: {
                'Content-Type': `multipart/mixed; boundary=batch_${batchId}`,
                'Accept': 'application/json;odata=verbose'
            },
            data: requestBody.join('')
        };
        const response = await request.post(requestOptions);
        const errors = this.parseBatchResponseBody(response, chunk);
        if (errors.length > 0) {
            throw `Creating some items failed with the following errors: ${os.EOL}${errors.map(error => { return `- ${error}`; }).join(os.EOL)}`;
        }
    }
    getRequestBody(chunk, batchId, options) {
        const changeSetId = v4();
        const batchBody = [];
        batchBody.push(`--batch_${batchId}\n`);
        batchBody.push(`Content-Type: multipart/mixed; boundary="changeset_${changeSetId}"\n\n`);
        batchBody.push('Content-Transfer-Encoding: binary\n\n');
        for (const item of chunk) {
            const actionUrl = this.getActionUrl(options, item);
            batchBody.push(`--changeset_${changeSetId}\n`);
            batchBody.push('Content-Type: application/http\n');
            batchBody.push('Content-Transfer-Encoding: binary\n\n');
            batchBody.push(`DELETE ${actionUrl} HTTP/1.1\n`);
            batchBody.push(`Accept: application/json;odata=nometadata\n`);
            batchBody.push(`If-Match: *\n\n`);
        }
        batchBody.push(`\n\n`);
        batchBody.push(`--changeset_${changeSetId}--\n\n`);
        batchBody.push(`--batch_${batchId}--\n`);
        return batchBody;
    }
    parseBatchResponseBody(response, chunk) {
        const errors = [];
        response.split('\r\n')
            .filter((line) => line.startsWith('{'))
            .forEach((line, index) => {
            const parsedResponse = JSON.parse(line);
            if (parsedResponse.error) {
                const error = parsedResponse.error;
                errors.push(`Item ID ${chunk[index]}: ${error.message.value}`);
            }
        });
        return errors;
    }
    ;
    getChunkedArray(inputArray, chunkSize) {
        const result = [];
        for (let i = 0; i < inputArray.length; i += chunkSize) {
            result.push(inputArray.slice(i, i + chunkSize));
        }
        return result;
    }
    getActionUrl(options, item) {
        let requestUrl = '';
        if (options.listId) {
            requestUrl += `lists(guid'${formatting.encodeQueryParameter(options.listId)}')`;
        }
        else if (options.listTitle) {
            requestUrl += `lists/getByTitle('${formatting.encodeQueryParameter(options.listTitle)}')`;
        }
        else if (options.listUrl) {
            const listServerRelativeUrl = urlUtil.getServerRelativePath(options.webUrl, options.listUrl);
            requestUrl += `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
        }
        requestUrl += `/items(${item})`;
        if (options.recycle) {
            requestUrl += '/recycle()';
        }
        return requestUrl;
    }
}
_SpoListItemBatchRemoveCommand_instances = new WeakSet(), _SpoListItemBatchRemoveCommand_initTelemetry = function _SpoListItemBatchRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            filePath: typeof args.options.filePath !== 'undefined',
            ids: typeof args.options.ids !== 'undefined',
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            recycle: !!args.options.recycle,
            force: !!args.options.force
        });
    });
}, _SpoListItemBatchRemoveCommand_initOptions = function _SpoListItemBatchRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-l, --listId [listId]'
    }, {
        option: '-t, --listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '-p, --filePath [filePath]'
    }, {
        option: '-i, --ids [ids]'
    }, {
        option: '-r, --recycle'
    }, {
        option: '-f, --force'
    });
}, _SpoListItemBatchRemoveCommand_initValidators = function _SpoListItemBatchRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.listId &&
            !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} in option listId is not a valid GUID.`;
        }
        if (args.options.filePath) {
            if (!fs.existsSync(args.options.filePath)) {
                return `File with path ${args.options.filePath} does not exist.`;
            }
            const fileContent = fs.readFileSync(args.options.filePath, 'utf-8');
            const jsonContent = formatting.parseCsvToJson(fileContent);
            const idKey = Object.keys(jsonContent[0]).find(key => key.toLowerCase() === 'id');
            if (!idKey) {
                return `The file does not contain the required header row with the column name 'ID'.`;
            }
            const invalidIDs = validation.isValidPositiveIntegerArray(jsonContent.map(element => element[idKey].toString().trim()).join(','));
            if (invalidIDs !== true) {
                return `The file contains one or more invalid IDs: '${invalidIDs}'.`;
            }
        }
        if (args.options.ids) {
            const isValidIntegerArray = validation.isValidPositiveIntegerArray(args.options.ids);
            if (isValidIntegerArray !== true) {
                return `Option 'ids' contains one or more invalid IDs: '${isValidIntegerArray}'.`;
            }
        }
        return true;
    });
}, _SpoListItemBatchRemoveCommand_initTypes = function _SpoListItemBatchRemoveCommand_initTypes() {
    this.types.string.push('webUrl', 'listId', 'listTitle', 'listUrl', 'ids', 'filePath');
}, _SpoListItemBatchRemoveCommand_initOptionSets = function _SpoListItemBatchRemoveCommand_initOptionSets() {
    this.optionSets.push({
        options: ['listId', 'listTitle', 'listUrl']
    }, {
        options: ['filePath', 'ids']
    });
};
export default new SpoListItemBatchRemoveCommand();
//# sourceMappingURL=listitem-batch-remove.js.map