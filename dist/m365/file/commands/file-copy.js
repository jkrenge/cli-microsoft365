var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _FileCopyCommand_instances, _FileCopyCommand_initTelemetry, _FileCopyCommand_initOptions, _FileCopyCommand_initValidators;
import request from '../../../request.js';
import { urlUtil } from '../../../utils/urlUtil.js';
import { spo } from '../../../utils/spo.js';
import { validation } from '../../../utils/validation.js';
import GraphCommand from '../../base/GraphCommand.js';
import commands from '../commands.js';
class FileCopyCommand extends GraphCommand {
    get name() {
        return commands.COPY;
    }
    get description() {
        return 'Copies a file to another location using the Microsoft Graph';
    }
    constructor() {
        super();
        _FileCopyCommand_instances.add(this);
        this.nameConflictBehaviorOptions = ['fail', 'replace', 'rename'];
        __classPrivateFieldGet(this, _FileCopyCommand_instances, "m", _FileCopyCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _FileCopyCommand_instances, "m", _FileCopyCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _FileCopyCommand_instances, "m", _FileCopyCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const { webUrl, sourceUrl, targetUrl, newName, verbose } = args.options;
            const sourcePath = this.getAbsoluteUrl(webUrl, sourceUrl);
            const destinationPath = this.getAbsoluteUrl(webUrl, targetUrl);
            if (this.verbose) {
                await logger.logToStderr(`Copying file '${sourcePath}' to '${destinationPath}'...`);
            }
            const copyUrl = await this.getCopyUrl(args.options, sourcePath, logger);
            const { targetDriveId, targetItemId } = await this.getTargetDriveAndItemId(webUrl, targetUrl, logger, verbose);
            const requestOptions = {
                url: copyUrl,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: {
                    parentReference: {
                        driveId: targetDriveId,
                        id: targetItemId
                    }
                }
            };
            if (newName) {
                const sourceFileName = sourcePath.substring(sourcePath.lastIndexOf('/') + 1);
                const sourceFileExtension = sourceFileName.includes('.') ? sourceFileName.substring(sourceFileName.lastIndexOf('.')) : '';
                const newNameExtension = newName.includes('.') ? newName.substring(newName.lastIndexOf('.')) : '';
                requestOptions.data.name = newNameExtension ? `${newName.replace(newNameExtension, "")}${sourceFileExtension}` : `${newName}${sourceFileExtension}`;
            }
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getCopyUrl(options, sourcePath, logger) {
        const { webUrl, sourceUrl, verbose, nameConflictBehavior } = options;
        const folderUrl = new URL(sourcePath);
        const siteId = await spo.getSiteIdByMSGraph(webUrl, logger, verbose);
        const drive = await this.getDocumentLibrary(siteId, folderUrl, sourceUrl, logger);
        const itemId = await this.getStartingFolderId(drive, folderUrl, logger);
        const queryParameters = nameConflictBehavior && nameConflictBehavior !== 'fail'
            ? `@microsoft.graph.conflictBehavior=${nameConflictBehavior}`
            : '';
        const copyUrl = `${this.resource}/v1.0/sites/${siteId}/drives/${drive.id}/items/${itemId}/copy${queryParameters ? `?${queryParameters}` : ''}`;
        return copyUrl;
    }
    async getTargetDriveAndItemId(webUrl, targetUrl, logger, verbose) {
        const targetSiteUrl = urlUtil.getTargetSiteAbsoluteUrl(webUrl, targetUrl);
        const targetSiteId = await spo.getSiteIdByMSGraph(targetSiteUrl, logger, verbose);
        const targetFolderUrl = new URL(this.getAbsoluteUrl(targetSiteUrl, targetUrl));
        const targetDrive = await this.getDocumentLibrary(targetSiteId, targetFolderUrl, targetUrl, logger);
        const targetDriveId = targetDrive.id;
        const targetItemId = await this.getStartingFolderId(targetDrive, targetFolderUrl, logger);
        return { targetDriveId, targetItemId };
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
        return drive;
    }
    async getStartingFolderId(documentLibrary, folderUrl, logger) {
        if (this.verbose) {
            await logger.logToStderr(`Getting starting folder id...`);
        }
        const documentLibraryRelativeFolderUrl = folderUrl.href.replace(new RegExp(`${documentLibrary.webUrl}`, 'i'), '').replace(/\/+$/, '');
        const requestOptions = {
            url: `${this.resource}/v1.0/drives/${documentLibrary.id}/root${documentLibraryRelativeFolderUrl ? `:${documentLibraryRelativeFolderUrl}` : ''}?$select=id`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const folder = await request.get(requestOptions);
        return folder?.id;
    }
    getAbsoluteUrl(webUrl, url) {
        return url.startsWith('https://') ? url : urlUtil.getAbsoluteUrl(webUrl, url);
    }
}
_FileCopyCommand_instances = new WeakSet(), _FileCopyCommand_initTelemetry = function _FileCopyCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            webUrl: typeof args.options.webUrl !== 'undefined',
            sourceUrl: typeof args.options.sourceUrl !== 'undefined',
            targetUrl: typeof args.options.targetUrl !== 'undefined',
            newName: typeof args.options.newName !== 'undefined',
            nameConflictBehavior: typeof args.options.nameConflictBehavior !== 'undefined'
        });
    });
}, _FileCopyCommand_initOptions = function _FileCopyCommand_initOptions() {
    this.options.unshift({ option: '-u, --webUrl <webUrl>' }, { option: '-s, --sourceUrl <sourceUrl>' }, { option: '-t, --targetUrl <targetUrl>' }, { option: '--newName [newName]' }, { option: '--nameConflictBehavior [nameConflictBehavior]', autocomplete: this.nameConflictBehaviorOptions });
}, _FileCopyCommand_initValidators = function _FileCopyCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.nameConflictBehavior && this.nameConflictBehaviorOptions.indexOf(args.options.nameConflictBehavior) === -1) {
            return `${args.options.nameConflictBehavior} is not a valid nameConflictBehavior value. Allowed values: ${this.nameConflictBehaviorOptions.join(', ')}.`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
};
export default new FileCopyCommand();
//# sourceMappingURL=file-copy.js.map