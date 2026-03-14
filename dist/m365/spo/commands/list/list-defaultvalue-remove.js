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
    fieldName: z.string(),
    folderUrl: z.string().optional(),
    force: z.boolean().optional().alias('f')
});
class SpoListDefaultValueRemoveCommand extends SpoCommand {
    get name() {
        return commands.LIST_DEFAULTVALUE_REMOVE;
    }
    get description() {
        return 'Removes a specific default column value for a specific document library';
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
        if (!args.options.force) {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove default column value '${args.options.fieldName}' from ${args.options.folderUrl ? `'${args.options.folderUrl}'` : 'the root of the list'}?` });
            if (!result) {
                return;
            }
        }
        try {
            if (this.verbose) {
                await logger.logToStderr(`Removing default column value '${args.options.fieldName}' from ${args.options.folderUrl ? `'${args.options.folderUrl}'` : 'the root of the list'}.`);
                await logger.logToStderr(`Getting server-relative URL of the list...`);
            }
            const listServerRelUrl = await this.getServerRelativeListUrl(args.options);
            let folderUrl = listServerRelUrl;
            if (args.options.folderUrl) {
                folderUrl = urlUtil.getServerRelativePath(args.options.webUrl, urlUtil.removeTrailingSlashes(args.options.folderUrl));
            }
            if (this.verbose) {
                await logger.logToStderr(`Getting default column values...`);
            }
            const defaultValuesXml = await this.getDefaultColumnValuesXml(args.options.webUrl, listServerRelUrl);
            const removeDefaultValueResult = this.removeFieldFromXml(defaultValuesXml, args.options.fieldName, folderUrl);
            if (!removeDefaultValueResult.isFieldFound) {
                throw `Default column value '${args.options.fieldName}' was not found.`;
            }
            if (this.verbose) {
                await logger.logToStderr(`Uploading default column values to list...`);
            }
            await this.uploadDefaultColumnValuesXml(args.options.webUrl, listServerRelUrl, removeDefaultValueResult.xml);
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
    removeFieldFromXml(xml, fieldName, folderUrl) {
        if (xml === null) {
            return { isFieldFound: false };
        }
        // Encode all spaces in the folder URL
        const encodedFolderUrl = folderUrl.replace(/ /g, '%20');
        const parser = new DOMParser();
        const doc = parser.parseFromString(xml, 'application/xml');
        const folderLinks = doc.getElementsByTagName('a');
        for (let i = 0; i < folderLinks.length; i++) {
            const folderNode = folderLinks[i];
            const folderNodeUrl = folderNode.getAttribute('href');
            if (encodedFolderUrl.toLowerCase() !== folderNodeUrl.toLowerCase()) {
                continue;
            }
            const defaultValues = folderNode.getElementsByTagName('DefaultValue');
            for (let j = 0; j < defaultValues.length; j++) {
                const defaultValueNode = defaultValues[j];
                const defaultValueNodeField = defaultValueNode.getAttribute('FieldName');
                if (defaultValueNodeField !== fieldName) {
                    continue;
                }
                if (folderNode.childNodes.length === 1) {
                    // No other default values found in the folder, let's remove the folder node
                    folderNode.parentNode.removeChild(folderNode);
                }
                else {
                    // Default value node found, let's remove it
                    folderNode.removeChild(defaultValueNode);
                }
                return { isFieldFound: true, xml: doc.toString() };
            }
        }
        return { isFieldFound: false };
    }
    async uploadDefaultColumnValuesXml(webUrl, listServerRelUrl, xml) {
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
export default new SpoListDefaultValueRemoveCommand();
//# sourceMappingURL=list-defaultvalue-remove.js.map