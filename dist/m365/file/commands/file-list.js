var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _FileListCommand_instances, _FileListCommand_initTelemetry, _FileListCommand_initOptions, _FileListCommand_initValidators;
import request from '../../../request.js';
import { formatting } from '../../../utils/formatting.js';
import { odata } from '../../../utils/odata.js';
import { validation } from '../../../utils/validation.js';
import GraphCommand from '../../base/GraphCommand.js';
import commands from '../commands.js';
class FileListCommand extends GraphCommand {
    get name() {
        return commands.LIST;
    }
    get description() {
        return 'Retrieves files from the specified folder and site';
    }
    defaultProperties() {
        return ['name', 'lastModifiedByUser'];
    }
    constructor() {
        super();
        _FileListCommand_instances.add(this);
        this.foldersToGetFilesFrom = [];
        __classPrivateFieldGet(this, _FileListCommand_instances, "m", _FileListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _FileListCommand_instances, "m", _FileListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _FileListCommand_instances, "m", _FileListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        let webUrl = args.options.webUrl;
        if (!webUrl.endsWith('/')) {
            webUrl += '/';
        }
        let folderUrlValue = args.options.folderUrl;
        if (folderUrlValue.endsWith('/')) {
            folderUrlValue = folderUrlValue.slice(0, -1);
        }
        const folderUrl = new URL(folderUrlValue, webUrl);
        try {
            const siteId = await this.getSiteId(args.options.webUrl, logger);
            const drive = await this.getDocumentLibrary(siteId, folderUrl, args.options.folderUrl, logger);
            const driveId = drive.id;
            const folderId = await this.getStartingFolderId(drive, folderUrl, logger);
            if (this.verbose) {
                await logger.logToStderr(`Loading folders to get files from...`);
            }
            // add the starting folder to the list of folders to get files from
            this.foldersToGetFilesFrom.push(folderId);
            await this.loadFoldersToGetFilesFrom(folderId, driveId, args.options.recursive);
            if (this.debug) {
                await logger.logToStderr(`Folders to get files from: ${this.foldersToGetFilesFrom.join(', ')}`);
            }
            const files = await this.loadFilesFromFolders(driveId, this.foldersToGetFilesFrom, logger);
            await logger.log(files);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getSiteId(webUrl, logger) {
        if (this.verbose) {
            await logger.logToStderr(`Getting site id...`);
        }
        const url = new URL(webUrl);
        const requestOptions = {
            url: `${this.resource}/v1.0/sites/${formatting.encodeQueryParameter(url.host)}:${url.pathname}?$select=id`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        return request
            .get(requestOptions)
            .then(async (site) => {
            if (this.verbose) {
                await logger.logToStderr(`Site id: ${site.id}`);
            }
            return site.id;
        });
    }
    async getDocumentLibrary(siteId, folderUrl, folderUrlFromUser, logger) {
        if (this.verbose) {
            await logger.logToStderr(`Getting document library...`);
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/sites/${siteId}/drives?$select=webUrl,id`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const drives = await request.get(requestOptions);
        const lowerCaseFolderUrl = folderUrl.href.toLowerCase();
        const drive = drives.value
            .sort((a, b) => b.webUrl.localeCompare(a.webUrl))
            .find((d) => {
            const driveUrl = d.webUrl.toLowerCase();
            // ensure that the drive url is a prefix of the folder url
            return lowerCaseFolderUrl.startsWith(driveUrl) &&
                (driveUrl.length === lowerCaseFolderUrl.length ||
                    lowerCaseFolderUrl[driveUrl.length] === '/');
        });
        if (!drive) {
            throw `Document library '${folderUrlFromUser}' not found`;
        }
        if (this.verbose) {
            await logger.logToStderr(`Document library: ${drive.webUrl}, ${drive.id}`);
        }
        return drive;
    }
    async getStartingFolderId(documentLibrary, folderUrl, logger) {
        if (this.verbose) {
            await logger.logToStderr(`Getting starting folder id...`);
        }
        const documentLibraryRelativeFolderUrl = folderUrl.href.replace(new RegExp(documentLibrary.webUrl, 'i'), '');
        const requestOptions = {
            url: `${this.resource}/v1.0/drives/${documentLibrary.id}/root${documentLibraryRelativeFolderUrl.length > 0 ? `:${documentLibraryRelativeFolderUrl}` : ''}?$select=id`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const folder = await request.get(requestOptions);
        if (this.verbose) {
            await logger.logToStderr(`Starting folder id: ${folder.id}`);
        }
        return folder.id;
    }
    async loadFoldersToGetFilesFrom(folderId, driveId, recursive) {
        if (!recursive) {
            return;
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/drives/${driveId}/items('${folderId}')/children?$filter=folder ne null&$select=id`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const subfolders = await request.get(requestOptions);
        const subfolderIds = subfolders.value.map((subfolder) => subfolder.id);
        this.foldersToGetFilesFrom = this.foldersToGetFilesFrom.concat(subfolderIds);
        await Promise.all(subfolderIds.map((subfolderId) => this.loadFoldersToGetFilesFrom(subfolderId, driveId, recursive)));
    }
    async loadFilesFromFolders(driveId, folderIds, logger) {
        if (this.verbose) {
            await logger.logToStderr(`Loading files from folders...`);
        }
        let files = [];
        const res = await Promise.all(folderIds.map((folderId) => 
        // get items from folder. Because we can't filter out folders here
        // we need to get all items from the folder and filter them out later
        odata.getAllItems(`${this.resource}/v1.0/drives/${driveId}/items/${folderId}/children`)));
        // flatten data from all promises
        files = files.concat(...res);
        // remove folders from the list of files
        files = files.filter((item) => item.file);
        files.forEach(file => file.lastModifiedByUser = file.lastModifiedBy?.user?.displayName);
        return files;
    }
}
_FileListCommand_instances = new WeakSet(), _FileListCommand_initTelemetry = function _FileListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            recursive: !!args.options.recursive
        });
    });
}, _FileListCommand_initOptions = function _FileListCommand_initOptions() {
    this.options.unshift({ option: '-u, --webUrl <webUrl>' }, { option: '--folderUrl <folderUrl>' }, { option: '--recursive' });
}, _FileListCommand_initValidators = function _FileListCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.webUrl));
};
export default new FileListCommand();
//# sourceMappingURL=file-list.js.map