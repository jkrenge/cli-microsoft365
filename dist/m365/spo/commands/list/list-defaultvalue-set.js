import SpoCommand from '../../../base/SpoCommand.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import commands from '../../commands.js';
import { DOMParser } from '@xmldom/xmldom';
import { validation } from '../../../../utils/validation.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
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
    fieldName: z.string(),
    fieldValue: z.string()
        .refine(value => value !== '', `The value cannot be empty. Use 'spo list defaultvalue remove' to remove a default column value.`),
    folderUrl: z.string().optional()
        .refine(url => url === undefined || (!url.includes('#') && !url.includes('%')), 'Due to limitations in SharePoint Online, setting default column values for folders with a # or % character in their path is not supported.')
});
class SpoListDefaultValueSetCommand extends SpoCommand {
    get name() {
        return commands.LIST_DEFAULTVALUE_SET;
    }
    get description() {
        return 'Sets default column values for a specific document library';
    }
    get schema() {
        return options;
    }
    getRefinedSchema(schema) {
        return schema
            .refine(options => [options.listId, options.listTitle, options.listUrl].filter(o => o !== undefined).length === 1, {
            error: 'Use one of the following options: listId, listTitle, listUrl.'
        });
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Setting default column value '${args.options.fieldValue}' for field '${args.options.fieldName}'...`);
                await logger.logToStderr(`Getting server-relative URL of the list...`);
            }
            const listServerRelUrl = await this.getServerRelativeListUrl(args.options);
            let folderUrl = listServerRelUrl;
            if (args.options.folderUrl) {
                if (this.verbose) {
                    await logger.logToStderr(`Getting server-relative URL of folder '${args.options.folderUrl}'...`);
                }
                // Casing of the folder URL is important, let's retrieve the correct URL
                const serverRelativeFolderUrl = urlUtil.getServerRelativePath(args.options.webUrl, urlUtil.removeTrailingSlashes(args.options.folderUrl));
                folderUrl = await this.getCorrectFolderUrl(args.options.webUrl, serverRelativeFolderUrl);
            }
            if (this.verbose) {
                await logger.logToStderr(`Getting default column values...`);
            }
            const defaultValuesXml = await this.ensureDefaultColumnValuesXml(args.options.webUrl, listServerRelUrl);
            const modifiedXml = await this.updateFieldValueXml(logger, defaultValuesXml, args.options.fieldName, args.options.fieldValue, folderUrl);
            await this.uploadDefaultColumnValuesXml(logger, args.options.webUrl, listServerRelUrl, modifiedXml);
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
    async getCorrectFolderUrl(webUrl, folderUrl) {
        const requestOptions = {
            // Using ListItemAllFields endpoint because GetFolderByServerRelativePath doesn't return the correctly cased URL
            url: `${webUrl}/_api/Web/GetFolderByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(folderUrl)}')/ListItemAllFields?$select=FileRef`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        if (!response.FileRef) {
            throw `Folder '${folderUrl}' was not found.`;
        }
        return response.FileRef;
    }
    async ensureDefaultColumnValuesXml(webUrl, listServerRelUrl) {
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
            if (err.status !== 404) {
                throw err;
            }
            // For lists that have never had default column values set, the client_LocationBasedDefaults.html file does not exist.
            // In this case, we need to create the file with blank default metadata.
            const requestOptions = {
                url: `${webUrl}/_api/Web/GetFolderByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listServerRelUrl + '/Forms')}')/Files/Add(url='client_LocationBasedDefaults.html', overwrite=false)`,
                headers: {
                    accept: 'application/json;odata=nometadata',
                    'content-type': 'text/plain'
                },
                responseType: 'json',
                data: '<MetadataDefaults />'
            };
            await request.post(requestOptions);
            return requestOptions.data;
        }
    }
    async updateFieldValueXml(logger, xml, fieldName, fieldValue, folderUrl) {
        if (this.verbose) {
            await logger.logToStderr(`Modifying default column values...`);
        }
        // Encode all spaces in the folder URL
        const encodedFolderUrl = folderUrl.replace(/ /g, '%20');
        const parser = new DOMParser();
        const doc = parser.parseFromString(xml, 'application/xml');
        // Create a new DefaultValue node
        const newDefaultValueNode = doc.createElement('DefaultValue');
        newDefaultValueNode.setAttribute('FieldName', fieldName);
        newDefaultValueNode.textContent = fieldValue;
        const folderLinks = doc.getElementsByTagName('a');
        for (let i = 0; i < folderLinks.length; i++) {
            const folderNode = folderLinks[i];
            const folderNodeUrl = folderNode.getAttribute('href');
            if (encodedFolderUrl !== folderNodeUrl) {
                continue;
            }
            const defaultValues = folderNode.getElementsByTagName('DefaultValue');
            for (let j = 0; j < defaultValues.length; j++) {
                const defaultValueNode = defaultValues[j];
                const defaultValueNodeField = defaultValueNode.getAttribute('FieldName');
                if (defaultValueNodeField !== fieldName) {
                    continue;
                }
                // Default value node found, let's update the value
                defaultValueNode.textContent = fieldValue;
                return doc.toString();
            }
            // Default value node not found, let's create it
            folderNode.appendChild(newDefaultValueNode);
            return doc.toString();
        }
        // Folder node was not found, let's create it
        const newFolderNode = doc.createElement('a');
        newFolderNode.setAttribute('href', encodedFolderUrl);
        newFolderNode.appendChild(newDefaultValueNode);
        doc.documentElement.appendChild(newFolderNode);
        return doc.toString();
    }
    async uploadDefaultColumnValuesXml(logger, webUrl, listServerRelUrl, xml) {
        if (this.verbose) {
            await logger.logToStderr(`Uploading default column values to list...`);
        }
        const requestOptions = {
            url: `${webUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listServerRelUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`,
            headers: {
                accept: 'application/json;odata=nometadata',
                'content-type': 'text/plain'
            },
            responseType: 'json',
            data: xml
        };
        await request.put(requestOptions);
    }
}
export default new SpoListDefaultValueSetCommand();
//# sourceMappingURL=list-defaultvalue-set.js.map