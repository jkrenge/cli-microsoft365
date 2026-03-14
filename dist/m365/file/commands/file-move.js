var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _FileMoveCommand_instances, _FileMoveCommand_initTelemetry, _FileMoveCommand_initOptions, _FileMoveCommand_initValidators;
import GraphCommand from '../../base/GraphCommand.js';
import { setTimeout } from 'timers/promises';
import commands from '../commands.js';
import request from '../../../request.js';
import { spo } from '../../../utils/spo.js';
import { urlUtil } from '../../../utils/urlUtil.js';
import { drive } from '../../../utils/drive.js';
import { validation } from '../../../utils/validation.js';
class FileMoveCommand extends GraphCommand {
    get name() {
        return commands.MOVE;
    }
    get description() {
        return 'Moves a file to another location using the Microsoft Graph';
    }
    constructor() {
        super();
        _FileMoveCommand_instances.add(this);
        this.pollingInterval = 10000;
        this.nameConflictBehaviorOptions = ['fail', 'replace', 'rename'];
        __classPrivateFieldGet(this, _FileMoveCommand_instances, "m", _FileMoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _FileMoveCommand_instances, "m", _FileMoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _FileMoveCommand_instances, "m", _FileMoveCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const { webUrl, sourceUrl, targetUrl, nameConflictBehavior, newName, verbose } = args.options;
            const sourcePath = this.getAbsoluteUrl(webUrl, sourceUrl);
            const destinationPath = this.getAbsoluteUrl(webUrl, targetUrl);
            const { driveId, itemId } = await this.getDriveIdAndItemId(webUrl, sourcePath, sourceUrl, logger, verbose);
            const targetSiteUrl = urlUtil.getTargetSiteAbsoluteUrl(webUrl, targetUrl);
            const targetFolderUrl = this.getAbsoluteUrl(targetSiteUrl, targetUrl);
            const { driveId: targetDriveId, itemId: targetItemId } = await this.getDriveIdAndItemId(targetSiteUrl, targetFolderUrl, targetUrl, logger, verbose);
            const requestOptions = this.getRequestOptions(driveId, itemId, targetDriveId, targetItemId, newName, sourcePath, nameConflictBehavior);
            if (verbose) {
                await logger.logToStderr(`Moving file '${sourcePath}' to '${destinationPath}'...`);
            }
            if (driveId === targetDriveId) {
                await request.patch(requestOptions);
            }
            else {
                const response = await request.post(requestOptions);
                await this.waitUntilCopyOperationCompleted(response.headers.location, logger);
                const itemUrl = `${this.resource}/v1.0/drives/${driveId}/items/${itemId}`;
                await request.delete({ url: itemUrl, headers: requestOptions.headers });
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getAbsoluteUrl(webUrl, url) {
        return url.startsWith('https://') ? url : urlUtil.getAbsoluteUrl(webUrl, url);
    }
    async getDriveIdAndItemId(webUrl, folderUrl, sourceUrl, logger, verbose) {
        const siteId = await spo.getSiteIdByMSGraph(webUrl, logger, verbose);
        const driveDetails = await drive.getDriveByUrl(siteId, new URL(folderUrl), logger, verbose);
        const itemId = await drive.getDriveItemId(driveDetails, new URL(folderUrl), logger, verbose);
        return { driveId: driveDetails.id, itemId };
    }
    getRequestOptions(sourceDriveId, sourceItemId, targetDriveId, targetItemId, newName, sourcePath, nameConflictBehavior) {
        const apiUrl = sourceDriveId === targetDriveId
            ? `${this.resource}/v1.0/drives/${sourceDriveId}/items/${sourceItemId}`
            : `${this.resource}/v1.0/drives/${sourceDriveId}/items/${sourceItemId}/copy`;
        const queryParameters = nameConflictBehavior && nameConflictBehavior !== 'fail'
            ? `@microsoft.graph.conflictBehavior=${nameConflictBehavior}`
            : '';
        const urlWithQuery = `${apiUrl}${queryParameters ? `?${queryParameters}` : ''}`;
        const requestOptions = {
            url: urlWithQuery,
            headers: { accept: 'application/json;odata.metadata=none' },
            responseType: 'json',
            fullResponse: true,
            data: { parentReference: { driveId: targetDriveId, id: targetItemId } }
        };
        if (newName) {
            const sourceFileName = sourcePath.substring(sourcePath.lastIndexOf('/') + 1);
            const sourceFileExtension = sourceFileName.includes('.') ? sourceFileName.substring(sourceFileName.lastIndexOf('.')) : '';
            const newNameExtension = newName.includes('.') ? newName.substring(newName.lastIndexOf('.')) : '';
            requestOptions.data.name = newNameExtension ? `${newName.replace(newNameExtension, "")}${sourceFileExtension}` : `${newName}${sourceFileExtension}`;
        }
        return requestOptions;
    }
    async waitUntilCopyOperationCompleted(monitorUrl, logger) {
        const requestOptions = {
            url: monitorUrl,
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        if (response.status === 'completed') {
            if (this.verbose) {
                await logger.logToStderr('Copy operation completed succesfully. Returning...');
            }
            return;
        }
        else if (response.status === 'failed') {
            throw response.error.message;
        }
        else {
            if (this.verbose) {
                await logger.logToStderr(`Still copying. Retrying in ${this.pollingInterval / 1000} seconds...`);
            }
            await setTimeout(this.pollingInterval);
            await this.waitUntilCopyOperationCompleted(monitorUrl, logger);
        }
    }
}
_FileMoveCommand_instances = new WeakSet(), _FileMoveCommand_initTelemetry = function _FileMoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            webUrl: typeof args.options.webUrl !== 'undefined',
            sourceUrl: typeof args.options.sourceUrl !== 'undefined',
            targetUrl: typeof args.options.targetUrl !== 'undefined',
            newName: typeof args.options.newName !== 'undefined',
            nameConflictBehavior: typeof args.options.nameConflictBehavior !== 'undefined'
        });
    });
}, _FileMoveCommand_initOptions = function _FileMoveCommand_initOptions() {
    this.options.unshift({ option: '-u, --webUrl <webUrl>' }, { option: '-s, --sourceUrl <sourceUrl>' }, { option: '-t, --targetUrl <targetUrl>' }, { option: '--newName [newName]' }, { option: '--nameConflictBehavior [nameConflictBehavior]', autocomplete: this.nameConflictBehaviorOptions });
}, _FileMoveCommand_initValidators = function _FileMoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.nameConflictBehavior && this.nameConflictBehaviorOptions.indexOf(args.options.nameConflictBehavior) === -1) {
            return `${args.options.nameConflictBehavior} is not a valid nameConflictBehavior value. Allowed values: ${this.nameConflictBehaviorOptions.join(', ')}.`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
};
export default new FileMoveCommand();
//# sourceMappingURL=file-move.js.map