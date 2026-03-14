import SpoCommand from '../../../base/SpoCommand.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import commands from '../../commands.js';
import { DOMParser } from '@xmldom/xmldom';
import { validation } from '../../../../utils/validation.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { cli } from '../../../../cli/cli.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    webUrl: z.string()
        .refine(url => validation.isValidSharePointUrl(url) === true, {
        error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
        .alias('u'),
    listId: z.uuid().optional().alias('i'),
    listTitle: z.string().optional().alias('t'),
    listUrl: z.string().optional(),
    fieldName: z.string().optional(),
    folderUrl: z.string().optional(),
    force: z.boolean().optional().alias('f')
});
class SpoListDefaultValueClearCommand extends SpoCommand {
    get name() {
        return commands.LIST_DEFAULTVALUE_CLEAR;
    }
    get description() {
        return 'Clears default column values for a specific document library';
    }
    get schema() {
        return options;
    }
    getRefinedSchema(schema) {
        return schema
            .refine(options => [options.listId, options.listTitle, options.listUrl].filter(o => o !== undefined).length === 1, {
            error: 'Use one of the following options: listId, listTitle, listUrl.'
        })
            .refine(options => (options.fieldName !== undefined) !== (options.folderUrl !== undefined) || (options.fieldName === undefined && options.folderUrl === undefined), {
            error: `Specify 'fieldName' or 'folderUrl', but not both.`
        });
    }
    async commandAction(logger, args) {
        if (!args.options.force) {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to clear all default values${args.options.fieldName ? ` for field '${args.options.fieldName}'` : args.options.folderUrl ? ` for folder ${args.options.folderUrl}` : ''}?` });
            if (!result) {
                return;
            }
        }
        try {
            if (this.verbose) {
                await logger.logToStderr(`Clearing all default column values${args.options.fieldName ? ` for field ${args.options.fieldName}` : args.options.folderUrl ? `for folder '${args.options.folderUrl}'` : ''}...`);
                await logger.logToStderr(`Getting server-relative URL of the list...`);
            }
            const listServerRelUrl = await this.getServerRelativeListUrl(args.options);
            if (this.verbose) {
                await logger.logToStderr(`List server-relative URL: ${listServerRelUrl}`);
                await logger.logToStderr(`Getting default column values...`);
            }
            const defaultValuesXml = await this.getDefaultColumnValuesXml(args.options.webUrl, listServerRelUrl);
            if (defaultValuesXml === null) {
                if (this.verbose) {
                    await logger.logToStderr(`No default column values found.`);
                }
                return;
            }
            const trimmedXml = this.removeFieldsFromXml(defaultValuesXml, args.options);
            await this.uploadDefaultColumnValuesXml(args.options.webUrl, listServerRelUrl, trimmedXml);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getServerRelativeListUrl(options) {
        const requestOptions = {
            url: `${options.webUrl}/_api/Web`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        if (options.listUrl) {
            const serverRelativeUrl = urlUtil.getServerRelativePath(options.webUrl, options.listUrl);
            requestOptions.url += `/GetList('${formatting.encodeQueryParameter(serverRelativeUrl)}')`;
        }
        else if (options.listId) {
            requestOptions.url += `/Lists('${options.listId}')`;
        }
        else if (options.listTitle) {
            requestOptions.url += `/Lists/GetByTitle('${formatting.encodeQueryParameter(options.listTitle)}')`;
        }
        requestOptions.url += '?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate';
        try {
            const response = await request.get(requestOptions);
            if (response.BaseTemplate !== 101) {
                throw `The specified list is not a document library.`;
            }
            return response.RootFolder.ServerRelativeUrl;
        }
        catch (error) {
            if (error.status === 404) {
                throw `List '${options.listId || options.listTitle || options.listUrl}' was not found.`;
            }
            throw error;
        }
    }
    async getDefaultColumnValuesXml(webUrl, listServerRelUrl) {
        try {
            const requestOptions = {
                url: `${webUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listServerRelUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`,
                headers: {
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            const defaultValuesXml = await request.get(requestOptions);
            return defaultValuesXml;
        }
        catch (err) {
            // For lists that have never had default column values set, the client_LocationBasedDefaults.html file does not exist.
            if (err.status === 404) {
                return null;
            }
            throw err;
        }
    }
    removeFieldsFromXml(xml, options) {
        if (!options.fieldName && !options.folderUrl) {
            return '<MetadataDefaults />';
        }
        let folderUrlToRemove = null;
        if (options.folderUrl) {
            folderUrlToRemove = urlUtil.removeTrailingSlashes(urlUtil.getServerRelativePath(options.webUrl, options.folderUrl));
        }
        const parser = new DOMParser();
        const doc = parser.parseFromString(xml, 'application/xml');
        const folderLinks = doc.getElementsByTagName('a');
        for (let i = 0; i < folderLinks.length; i++) {
            const folderNode = folderLinks[i];
            const folderUrl = folderNode.getAttribute('href');
            if (folderUrlToRemove && folderUrlToRemove.toLowerCase() === decodeURIComponent(folderUrl).toLowerCase()) {
                folderNode.parentNode.removeChild(folderNode);
                break;
            }
            else if (options.fieldName) {
                const defaultValues = folderNode.getElementsByTagName('DefaultValue');
                for (let j = 0; j < defaultValues.length; j++) {
                    const defaultValueNode = defaultValues[j];
                    const fieldName = defaultValueNode.getAttribute('FieldName');
                    if (fieldName.toLowerCase() === options.fieldName.toLowerCase()) {
                        // Remove the entire folder node if it becomes empty
                        if (folderNode.childNodes.length === 1) {
                            folderNode.parentNode.removeChild(defaultValueNode.parentNode);
                        }
                        else {
                            folderNode.removeChild(defaultValueNode);
                        }
                        break;
                    }
                }
            }
        }
        return doc.toString();
    }
    async uploadDefaultColumnValuesXml(webUrl, listServerRelUrl, xml) {
        const requestOptions = {
            url: `${webUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listServerRelUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`,
            headers: {
                accept: 'application/json;odata=nometadata',
                'If-Match': '*'
            },
            responseType: 'json',
            data: xml
        };
        await request.put(requestOptions);
    }
}
export default new SpoListDefaultValueClearCommand();
//# sourceMappingURL=list-defaultvalue-clear.js.map