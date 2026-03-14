import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { globalOptionsZod } from '../../../../Command.js';
import z from 'zod';
var AllowedFieldTypeKind;
(function (AllowedFieldTypeKind) {
    AllowedFieldTypeKind[AllowedFieldTypeKind["Integer"] = 1] = "Integer";
    AllowedFieldTypeKind[AllowedFieldTypeKind["Text"] = 2] = "Text";
    AllowedFieldTypeKind[AllowedFieldTypeKind["Note"] = 3] = "Note";
    AllowedFieldTypeKind[AllowedFieldTypeKind["DateTime"] = 4] = "DateTime";
    AllowedFieldTypeKind[AllowedFieldTypeKind["Counter"] = 5] = "Counter";
    AllowedFieldTypeKind[AllowedFieldTypeKind["Choice"] = 6] = "Choice";
    AllowedFieldTypeKind[AllowedFieldTypeKind["Boolean"] = 8] = "Boolean";
    AllowedFieldTypeKind[AllowedFieldTypeKind["Number"] = 9] = "Number";
    AllowedFieldTypeKind[AllowedFieldTypeKind["Currency"] = 10] = "Currency";
    AllowedFieldTypeKind[AllowedFieldTypeKind["URL"] = 11] = "URL";
    AllowedFieldTypeKind[AllowedFieldTypeKind["Computed"] = 12] = "Computed";
    AllowedFieldTypeKind[AllowedFieldTypeKind["MultiChoice"] = 15] = "MultiChoice";
    AllowedFieldTypeKind[AllowedFieldTypeKind["GridChoice"] = 16] = "GridChoice";
})(AllowedFieldTypeKind || (AllowedFieldTypeKind = {}));
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    siteUrl: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
        error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    }).alias('u'),
    listTitle: z.string().optional(),
    listId: z.string().uuid()
        .refine(value => validation.isValidGuid(value), {
        error: e => `'${e.input}' in parameter listId is not a valid GUID.`
    }).optional(),
    listUrl: z.string().optional(),
    columnId: z.string().uuid()
        .refine(value => validation.isValidGuid(value), {
        error: e => `'${e.input}' in parameter columnId is not a valid GUID.`
    }).optional().alias('i'),
    columnTitle: z.string().optional().alias('t'),
    columnInternalName: z.string().optional(),
    prompt: z.string().optional(),
    isEnabled: z.boolean().optional()
}).strict();
class SppAutofillColumnSetCommand extends SpoCommand {
    get name() {
        return commands.AUTOFILLCOLUMN_SET;
    }
    get description() {
        return 'Applies the autofill option to the selected column';
    }
    get schema() {
        return options;
    }
    getRefinedSchema(schema) {
        return schema
            .refine(options => [options.columnId, options.columnTitle, options.columnInternalName].filter(Boolean).length === 1, {
            message: `Specify exactly one of the following options: 'columnId', 'columnTitle' or 'columnInternalName'.`
        })
            .refine(options => [options.listTitle, options.listId, options.listUrl].filter(Boolean).length === 1, {
            message: `Specify exactly one of the following options: 'listTitle', 'listId' or 'listUrl'.`
        });
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.log(`Applying an autofill column to a column...`);
            }
            const siteUrl = urlUtil.removeTrailingSlashes(args.options.siteUrl);
            const listInstance = await this.getDocumentLibraryInfo(siteUrl, args.options);
            if (listInstance.BaseType !== 1) {
                throw Error(`The specified list is not a document library.`);
            }
            const column = await this.getColumn(siteUrl, args.options, listInstance.Id);
            if (!Object.values(AllowedFieldTypeKind).includes(column.FieldTypeKind)) {
                throw Error(`The specified column has incorrect type.`);
            }
            if (column.AutofillInfo) {
                await this.updateAutoFillColumnSettings(siteUrl, args.options, column.Id, listInstance.Id, column.AutofillInfo);
                return;
            }
            if (!args.options.prompt) {
                throw Error(`The prompt parameter is required when setting the autofill column for the first time.`);
            }
            await this.applyAutoFillColumnSettings(siteUrl, args.options, column.Id, column.Title, listInstance.Id);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getDocumentLibraryInfo(siteUrl, options) {
        let requestUrl = `${siteUrl}/_api/web`;
        if (options.listId) {
            requestUrl += `/lists(guid'${formatting.encodeQueryParameter(options.listId)}')`;
        }
        else if (options.listTitle) {
            requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(options.listTitle)}')`;
        }
        else if (options.listUrl) {
            const listServerRelativeUrl = urlUtil.getServerRelativePath(siteUrl, options.listUrl);
            requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
        }
        const requestOptions = {
            url: `${requestUrl}?$select=Id,BaseType`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        return request.get(requestOptions);
    }
    getColumn(siteUrl, options, listId) {
        let fieldRestUrl;
        if (options.columnId) {
            fieldRestUrl = `/getbyid('${formatting.encodeQueryParameter(options.columnId)}')`;
        }
        else {
            fieldRestUrl = `/getbyinternalnameortitle('${formatting.encodeQueryParameter((options.columnTitle || options.columnInternalName))}')`;
        }
        const requestOptions = {
            url: `${siteUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(listId)}')/fields${fieldRestUrl}?&$select=Id,Title,FieldTypeKind,AutofillInfo`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        return request.get(requestOptions);
    }
    updateAutoFillColumnSettings(siteUrl, options, columnId, listInstanceId, autofillInfo) {
        const autofillInfoObj = JSON.parse(autofillInfo);
        const requestOptions = {
            url: `${siteUrl}/_api/machinelearning/SetColumnLLMInfo`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json',
            data: {
                autofillPrompt: options.prompt ?? autofillInfoObj.LLM.Prompt,
                columnId: columnId,
                docLibId: `{${listInstanceId}}`,
                isEnabled: options.isEnabled !== undefined ? options.isEnabled : autofillInfoObj.LLM.IsEnabled
            }
        };
        return request.post(requestOptions);
    }
    applyAutoFillColumnSettings(siteUrl, options, columnId, columnTitle, listInstanceId) {
        const requestOptions = {
            url: `${siteUrl}/_api/machinelearning/SetSyntexPoweredColumnPrompts`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            data: {
                docLibId: `{${listInstanceId}}`,
                syntexPoweredColumnPrompts: JSON.stringify([{
                        columnId: columnId,
                        columnName: columnTitle,
                        prompt: options.prompt,
                        isEnabled: options.isEnabled !== undefined ? options.isEnabled : true
                    }])
            }
        };
        return request.post(requestOptions);
    }
}
export default new SppAutofillColumnSetCommand();
//# sourceMappingURL=autofillcolumn-set.js.map