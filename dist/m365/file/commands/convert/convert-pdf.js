var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _FileConvertPdfCommand_instances, _FileConvertPdfCommand_initOptions, _FileConvertPdfCommand_initValidators;
import fs from 'fs';
import os from 'os';
import path from 'path';
import { v4 } from 'uuid';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class FileConvertPdfCommand extends GraphCommand {
    get name() {
        return commands.CONVERT_PDF;
    }
    get description() {
        return 'Converts the specified file to PDF using Microsoft Graph';
    }
    constructor() {
        super();
        _FileConvertPdfCommand_instances.add(this);
        __classPrivateFieldGet(this, _FileConvertPdfCommand_instances, "m", _FileConvertPdfCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _FileConvertPdfCommand_instances, "m", _FileConvertPdfCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        let sourceFileUrl = '';
        // path to the local file that contains the PDF-converted source file
        let localTargetFilePath = args.options.targetFile;
        let sourceIsLocalFile = true;
        let targetIsLocalFile = true;
        let error;
        const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken);
        if (typeof isAppOnlyAccessToken === 'undefined') {
            throw 'Unable to determine authentication type';
        }
        if (args.options.sourceFile.toLowerCase().startsWith('https://')) {
            sourceIsLocalFile = false;
        }
        if (args.options.targetFile.toLowerCase().startsWith('https://')) {
            localTargetFilePath = path.join(os.tmpdir(), v4());
            targetIsLocalFile = false;
            if (this.debug) {
                await logger.logToStderr(`Target set to a URL. Will store the temporary converted file at ${localTargetFilePath}`);
            }
        }
        try {
            try {
                sourceFileUrl = await this.getSourceFileUrl(logger, args, isAppOnlyAccessToken);
                const graphFileUrl = await this.getGraphFileUrl(logger, sourceFileUrl, this.sourceFileGraphUrl);
                const fileResponse = await this.convertFile(logger, graphFileUrl);
                await this.writeFileToDisk(logger, fileResponse, localTargetFilePath);
                await this.uploadConvertedFileIfNecessary(logger, targetIsLocalFile, localTargetFilePath, args.options.targetFile);
            }
            catch (err) {
                // catch the error from any of the previous promises so that we can
                // clean up resources in case something went wrong
                // if this.deleteRemoteSourceFileIfNecessary fails, it won't be caught
                // here, but rather at the end
                error = err;
            }
            await this.deleteRemoteSourceFileIfNecessary(logger, sourceIsLocalFile, sourceFileUrl);
            // if the target was a remote file, delete the local temp file
            if (!targetIsLocalFile) {
                if (this.verbose) {
                    await logger.logToStderr(`Deleting the temporary PDF file at ${localTargetFilePath}...`);
                }
                fs.unlinkSync(localTargetFilePath);
            }
            else {
                if (this.debug) {
                    await logger.logToStderr(`Target is a local path. Not deleting`);
                }
            }
            if (error) {
                this.handleRejectedODataJsonPromise(error);
            }
        }
        catch (err) {
            if (err instanceof CommandError) {
                throw err;
            }
            this.handleRejectedODataJsonPromise(err);
        }
    }
    /**
     * Returns web URL of the file to convert to PDF. If the user specified a URL
     * in command's options, returns the specified URL. If the user specified
     * a local file, it will upload the file and return its web URL. If CLI
     * is authenticated as app-only, uploads the file to the default document
     * library in the root site. If the CLI is authenticated as user, uploads the
     * file to the user's OneDrive for Business
     * @param await logger.logger instance
     * @param args Command args
     * @param isAppOnlyAccessToken True if CLI is authenticated in app-only mode
     * @returns Web URL of the file to upload
     */
    async getSourceFileUrl(logger, args, isAppOnlyAccessToken) {
        if (args.options.sourceFile.toLowerCase().startsWith('https://')) {
            return args.options.sourceFile;
        }
        if (this.verbose) {
            await logger.logToStderr('Uploading local file temporarily for conversion...');
        }
        const driveUrl = `${this.resource}/v1.0/${isAppOnlyAccessToken ? 'drive/root' : 'me/drive/root'}`;
        // we need the original file extension because otherwise Graph won't be able
        // to convert the file to PDF
        this.sourceFileGraphUrl = `${driveUrl}:/${v4()}${path.extname(args.options.sourceFile)}`;
        if (this.debug) {
            await logger.logToStderr(`Source is a local file. Uploading to ${this.sourceFileGraphUrl}...`);
        }
        return await this.uploadFile(args.options.sourceFile, this.sourceFileGraphUrl);
    }
    /**
     * Uploads the specified local file to a document library using Microsoft Graph
     * @param localFilePath Path to the local file to upload
     * @param targetGraphFileUrl Graph drive item URL of the file to upload
     * @returns Absolute URL of the uploaded file
     */
    async uploadFile(localFilePath, targetGraphFileUrl) {
        const requestOptions = {
            url: `${targetGraphFileUrl}:/createUploadSession`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const res = await request.post(requestOptions);
        const fileContents = fs.readFileSync(localFilePath);
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
     * @param fileGraphUrl If set, will return this URL without further action
     * @returns Graph's drive item URL for the specified file
     */
    async getGraphFileUrl(logger, fileWebUrl, fileGraphUrl) {
        if (this.debug) {
            await logger.logToStderr(`Resolving Graph drive item URL for ${fileWebUrl}`);
        }
        if (fileGraphUrl) {
            if (this.debug) {
                await logger.logToStderr(`Returning previously resolved Graph drive item URL ${fileGraphUrl}`);
            }
            return fileGraphUrl;
        }
        const _url = new URL(fileWebUrl);
        const siteInfo = await this.getGraphSiteInfoFromFullUrl(_url.hostname, _url.pathname);
        const siteId = siteInfo.id;
        let siteRelativeFileUrl = _url.pathname.replace(siteInfo.serverRelativeUrl, '');
        // normalize site-relative URLs for root site collections and root sites
        if (!siteRelativeFileUrl.startsWith('/')) {
            siteRelativeFileUrl = '/' + siteRelativeFileUrl;
        }
        const siteRelativeFileUrlChunks = siteRelativeFileUrl.split('/');
        const driveRelativeFileUrl = `/${siteRelativeFileUrlChunks.slice(2).join('/')}`;
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
     * @returns ID and server-relative URL of the site denoted by urlPath
     */
    async getGraphSiteInfoFromFullUrl(hostName, urlPath) {
        const siteId = '';
        const urlChunks = urlPath.split('/');
        return await this.getGraphSiteInfo(hostName, urlChunks, 0, siteId);
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
     * @param resolve Callback method to call when resolving site info succeeded
     * @param reject Callback method to call when resolving site info failed
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
            return await this.getGraphSiteInfo(hostName, urlChunks, ++currentChunk, res.id);
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
    /**
     * Requests conversion of a file to PDF using Microsoft Graph
     * @param await logger.logger instance
     * @param graphFileUrl Graph drive item URL of the file to convert to PDF
     * @returns Response object with a URL in the Location header that contains
     * the file converted to PDF. The URL must be called anonymously
     */
    async convertFile(logger, graphFileUrl) {
        if (this.verbose) {
            await logger.logToStderr('Converting file...');
        }
        const requestOptions = {
            url: `${graphFileUrl}:/content?format=pdf`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'stream'
        };
        return await request.get(requestOptions);
    }
    /**
     * Writes the contents of the specified file stream to a local file
     * @param await logger.logger instance
     * @param fileResponse Response with stream file contents
     * @param localFilePath Local file path where to store the file
     */
    async writeFileToDisk(logger, fileResponse, localFilePath) {
        if (this.verbose) {
            await logger.logToStderr(`Writing converted PDF file to ${localFilePath}...`);
        }
        await new Promise((resolve, reject) => {
            const writer = fs.createWriteStream(localFilePath);
            fileResponse.data.pipe(writer);
            writer.on('error', err => {
                reject(err);
            });
            writer.on('close', () => {
                resolve();
            });
        });
    }
    /**
     * If the user specified a URL as the targetFile, uploads the converted PDF
     * file to the specified location. If targetFile is a local path, doesn't do
     * anything.
     * @param await logger.logger instance
     * @param targetIsLocalFile Boolean that denotes if user specified as the target location a local path
     * @param localFilePath Local file path to where the file to be uploaded is located
     * @param targetFileUrl Web URL of the file to upload
     */
    async uploadConvertedFileIfNecessary(logger, targetIsLocalFile, localFilePath, targetFileUrl) {
        // if the target was a local path, we're done.
        // Otherwise, upload the file to the specified URL
        if (targetIsLocalFile) {
            if (this.debug) {
                await logger.logToStderr('Specified target is a local file. Not uploading.');
            }
            return;
        }
        if (this.verbose) {
            await logger.logToStderr(`Uploading converted PDF file to ${targetFileUrl}...`);
        }
        const targetGraphFileUrl = await this.getGraphFileUrl(logger, targetFileUrl, undefined);
        await this.uploadFile(localFilePath, targetGraphFileUrl);
    }
    /**
     * If the user specified local file to be converted to PDF, removes the file
     * that was temporarily upload to a document library for the conversion.
     * If the specified source file was a URL, doesn't do anything.
     * @param await logger.logger instance
     * @param sourceIsLocalFile Boolean that denotes if user specified a local path as the source file
     * @param sourceFileUrl Web URL of the temporary source file to delete
     */
    async deleteRemoteSourceFileIfNecessary(logger, sourceIsLocalFile, sourceFileUrl) {
        // if the source was a remote file, we're done,
        // otherwise delete the temporary uploaded file
        if (!sourceIsLocalFile) {
            if (this.debug) {
                await logger.logToStderr('Source file was URL. Not removing.');
            }
            return;
        }
        if (this.verbose) {
            await logger.logToStderr(`Deleting the temporary file at ${sourceFileUrl}...`);
        }
        const graphFileUrl = await this.getGraphFileUrl(logger, sourceFileUrl, this.sourceFileGraphUrl);
        const requestOptions = {
            url: graphFileUrl,
            headers: {
                accept: 'application/json;odata.metadata=none'
            }
        };
        return request.delete(requestOptions);
    }
}
_FileConvertPdfCommand_instances = new WeakSet(), _FileConvertPdfCommand_initOptions = function _FileConvertPdfCommand_initOptions() {
    this.options.unshift({
        option: '-s, --sourceFile <sourceFile>'
    }, {
        option: '-t, --targetFile <targetFile>'
    });
}, _FileConvertPdfCommand_initValidators = function _FileConvertPdfCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!args.options.sourceFile.toLowerCase().startsWith('https://') &&
            !fs.existsSync(args.options.sourceFile)) {
            // assume local path
            return `Specified source file ${args.options.sourceFile} doesn't exist`;
        }
        if (!args.options.targetFile.toLowerCase().startsWith('https://') &&
            fs.existsSync(args.options.targetFile)) {
            // assume local path
            return `Another file found at ${args.options.targetFile}`;
        }
        return true;
    });
};
export default new FileConvertPdfCommand();
//# sourceMappingURL=convert-pdf.js.map