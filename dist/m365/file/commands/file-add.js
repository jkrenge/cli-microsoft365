var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _FileAddCommand_instances, _FileAddCommand_initTelemetry, _FileAddCommand_initOptions, _FileAddCommand_initValidators;
import fs from 'fs';
import path from 'path';
import request from '../../../request.js';
import { validation } from '../../../utils/validation.js';
import GraphCommand from '../../base/GraphCommand.js';
import commands from '../commands.js';
class FileAddCommand extends GraphCommand {
    get name() {
        return commands.ADD;
    }
    get description() {
        return 'Uploads file to the specified site';
    }
    constructor() {
        super();
        _FileAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _FileAddCommand_instances, "m", _FileAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _FileAddCommand_instances, "m", _FileAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _FileAddCommand_instances, "m", _FileAddCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        let folderUrlWithoutTrailingSlash = args.options.folderUrl;
        if (folderUrlWithoutTrailingSlash.endsWith('/')) {
            folderUrlWithoutTrailingSlash = folderUrlWithoutTrailingSlash.substring(0, folderUrlWithoutTrailingSlash.length - 1);
        }
        try {
            const graphFileUrl = await this.getGraphFileUrl(logger, `${folderUrlWithoutTrailingSlash}/${path.basename(args.options.filePath)}`, args.options.siteUrl);
            await this.uploadFile(args.options.filePath, graphFileUrl);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    /**
     * Uploads the specified local file to a document library using Microsoft Graph
     * @param localFilePath Path to the local file to upload
     * @param targetGraphFileUrl Graph drive item URL of the file to upload
     * @returns Absolute URL of the uploaded file
     */
    async uploadFile(localFilePath, targetGraphFileUrl) {
        const fileContents = fs.readFileSync(localFilePath);
        const isEmptyFile = fileContents.length === 0;
        const requestOptions = {
            url: isEmptyFile ? `${targetGraphFileUrl}:/content` : `${targetGraphFileUrl}:/createUploadSession`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        if (isEmptyFile) {
            const res = await request.put(requestOptions);
            return res.webUrl;
        }
        else {
            const res = await request.post(requestOptions);
            const requestOptionsPut = {
                url: res.uploadUrl,
                headers: {
                    'x-anonymous': true,
                    'accept': 'application/json;odata.metadata=none',
                    'Content-Length': fileContents.length,
                    'Content-Range': `bytes 0-${fileContents.length - 1}/${fileContents.length}`
                },
                data: fileContents,
                responseType: 'json'
            };
            const resPut = await request.put(requestOptionsPut);
            return resPut.webUrl;
        }
    }
    /**
     * Gets Graph's drive item URL for the specified file. If the user specified
     * a local file to convert to PDF, returns the URL resolved while uploading
     * the file
     *
     * Example:
     *
     * fileWebUrl:
     * https://contoso.sharepoint.com/sites/Contoso/site/Shared%20Documents/file.docx
     *
     * returns:
     * https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,9d1b2174-9906-43ec-8c9e-f8589de047af,f60c833e-71ce-4a5a-b90e-2a7fdb718397/drives/b!k6NJ6ubjYEehsullOeFTcuYME3w1S8xHoHziURdWlu-DWrqz1yBLQI7E7_4TN6fL/root:/file.docx
     *
     * @param await logger.logger instance
     * @param fileWebUrl Web URL of the file for which to get drive item URL
     * @param siteUrl URL of the site to which upload the file. Optional. Specify to suppress lookup.
     * @returns Graph's drive item URL for the specified file
     */
    async getGraphFileUrl(logger, fileWebUrl, siteUrl) {
        if (this.debug) {
            await logger.logToStderr(`Resolving Graph drive item URL for ${fileWebUrl}`);
        }
        const _fileWebUrl = new URL(fileWebUrl);
        const _siteUrl = new URL(siteUrl || fileWebUrl);
        const isSiteUrl = typeof siteUrl !== 'undefined';
        const siteInfo = await this.getGraphSiteInfoFromFullUrl(_siteUrl.host, _siteUrl.pathname, isSiteUrl);
        let siteId = siteInfo.id;
        let siteRelativeFileUrl = _fileWebUrl.pathname.replace(siteInfo.serverRelativeUrl, '');
        // normalize site-relative URLs for root site collections and root sites
        if (!siteRelativeFileUrl.startsWith('/')) {
            siteRelativeFileUrl = '/' + siteRelativeFileUrl;
        }
        const siteRelativeFileUrlChunks = siteRelativeFileUrl.split('/');
        let driveRelativeFileUrl = `/${siteRelativeFileUrlChunks.slice(2).join('/')}`;
        // chunk 0 is empty because the URL starts with /
        const driveId = await this.getDriveId(logger, siteId, siteRelativeFileUrlChunks[1]);
        const graphUrl = `${this.resource}/v1.0/sites/${siteId}/drives/${driveId}/root:${driveRelativeFileUrl}`;
        if (this.debug) {
            await logger.logToStderr(`Resolved URL ${graphUrl}`);
        }
        return graphUrl;
    }
    /**
     * Retrieves the Graph ID and server-relative URL of the specified (sub)site.
     * Automatically detects which path chunks correspond to (sub)site.
     * @param hostName SharePoint host name, eg. contoso.sharepoint.com
     * @param urlPath Server-relative file URL, eg. /sites/site/docs/file1.aspx
     * @param isSiteUrl Set to true to indicate that the specified URL is a site URL
     * @returns ID and server-relative URL of the site denoted by urlPath
     */
    async getGraphSiteInfoFromFullUrl(hostName, urlPath, isSiteUrl) {
        const siteId = '';
        const urlChunks = urlPath.split('/');
        return await this.getGraphSiteInfo(hostName, urlChunks, isSiteUrl ? urlChunks.length - 1 : 0, siteId);
    }
    /**
     * Retrieves Graph site ID and server-relative URL of the site specified
     * using chunks from the URL path. Method is being called recursively as long
     * as it can successfully retrieve the site. When retrieving site fails, method
     * will return the last resolved site ID. If no site ID has been retrieved
     * (method fails on the first execution), it will call the reject callback.
     * @param hostName SharePoint host name, eg. contoso.sharepoint.com
     * @param urlChunks Array of chunks from server-relative URL, eg. ['sites', 'site', 'subsite', 'docs', 'file1.aspx']
     * @param currentChunk Current chunk that's being tested, eg. sites
     * @param lastSiteId Last correctly resolved Graph site ID
     * @returns Graph site ID and server-relative URL of the site specified through chunks
     */
    async getGraphSiteInfo(hostName, urlChunks, currentChunk, lastSiteId) {
        let currentPath = urlChunks.slice(0, currentChunk + 1).join('/');
        if (currentPath.endsWith('/sites') ||
            currentPath.endsWith('/teams') ||
            currentPath.endsWith('/personal')) {
            return await this.getGraphSiteInfo(hostName, urlChunks, currentChunk + 1, '');
        }
        if (!currentPath.startsWith('/')) {
            currentPath = '/' + currentPath;
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/sites/${hostName}:${currentPath}?$select=id`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const getResult = (id, serverRelativeUrl) => {
            return {
                id,
                serverRelativeUrl
            };
        };
        try {
            const res = await request.get(requestOptions);
            if (currentChunk === urlChunks.length - 1) {
                return getResult(res.id, currentPath);
            }
            else {
                return await this.getGraphSiteInfo(hostName, urlChunks, ++currentChunk, res.id);
            }
        }
        catch (err) {
            if (lastSiteId) {
                let serverRelativeUrl = `${urlChunks.slice(0, currentChunk).join('/')}`;
                if (!serverRelativeUrl.startsWith('/')) {
                    serverRelativeUrl = '/' + serverRelativeUrl;
                }
                return getResult(lastSiteId, serverRelativeUrl);
            }
            else {
                throw err;
            }
        }
    }
    /**
     * Returns the Graph drive ID of the specified document library
     * @param graphSiteId Graph ID of the site where the document library is located
     * @param siteRelativeListUrl Server-relative URL of the document library, eg. /sites/site/Documents
     * @returns Graph drive ID of the specified document library
     */
    async getDriveId(logger, graphSiteId, siteRelativeListUrl) {
        const requestOptions = {
            url: `${this.resource}/v1.0/sites/${graphSiteId}/drives?$select=webUrl,id`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const res = await request.get(requestOptions);
        if (this.debug) {
            await logger.logToStderr(`Searching for drive with a URL ending with /${siteRelativeListUrl}...`);
        }
        const drive = res.value.find(d => d.webUrl.endsWith(`/${siteRelativeListUrl}`));
        if (!drive) {
            throw 'Drive not found';
        }
        return drive.id;
    }
}
_FileAddCommand_instances = new WeakSet(), _FileAddCommand_initTelemetry = function _FileAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            siteUrl: typeof args.options.siteUrl !== 'undefined'
        });
    });
}, _FileAddCommand_initOptions = function _FileAddCommand_initOptions() {
    this.options.unshift({ option: '-u, --folderUrl <folderUrl>' }, { option: '-p, --filePath <filePath>' }, { option: '--siteUrl [siteUrl]' });
}, _FileAddCommand_initValidators = function _FileAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!fs.existsSync(args.options.filePath)) {
            return `Specified source file ${args.options.sourceFile} doesn't exist`;
        }
        if (args.options.siteUrl) {
            const isValidSiteUrl = validation.isValidSharePointUrl(args.options.siteUrl);
            if (isValidSiteUrl !== true) {
                return isValidSiteUrl;
            }
        }
        return validation.isValidSharePointUrl(args.options.folderUrl);
    });
};
export default new FileAddCommand();
//# sourceMappingURL=file-add.js.map