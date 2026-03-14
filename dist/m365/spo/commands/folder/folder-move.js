var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFolderMoveCommand_instances, _SpoFolderMoveCommand_initTelemetry, _SpoFolderMoveCommand_initOptions, _SpoFolderMoveCommand_initValidators, _SpoFolderMoveCommand_initOptionSets, _SpoFolderMoveCommand_initTypes;
import request from '../../../../request.js';
import { CreateFolderCopyJobsNameConflictBehavior, spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFolderMoveCommand extends SpoCommand {
    get name() {
        return commands.FOLDER_MOVE;
    }
    get description() {
        return 'Moves a folder to another location';
    }
    constructor() {
        super();
        _SpoFolderMoveCommand_instances.add(this);
        this.nameConflictBehaviorOptions = ['fail', 'rename'];
        __classPrivateFieldGet(this, _SpoFolderMoveCommand_instances, "m", _SpoFolderMoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFolderMoveCommand_instances, "m", _SpoFolderMoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFolderMoveCommand_instances, "m", _SpoFolderMoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFolderMoveCommand_instances, "m", _SpoFolderMoveCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFolderMoveCommand_instances, "m", _SpoFolderMoveCommand_initTypes).call(this);
    }
    getExcludedOptionsWithUrls() {
        return ['targetUrl', 'sourceUrl'];
    }
    async commandAction(logger, args) {
        try {
            const sourceServerRelativePath = await this.getSourcePath(logger, args.options);
            const sourcePath = this.getAbsoluteUrl(args.options.webUrl, sourceServerRelativePath);
            const destinationPath = this.getAbsoluteUrl(args.options.webUrl, args.options.targetUrl);
            if (this.verbose) {
                await logger.logToStderr(`Moving folder '${sourcePath}' to '${destinationPath}'...`);
            }
            const copyJobResponse = await spo.createFolderCopyJob(args.options.webUrl, sourcePath, destinationPath, {
                nameConflictBehavior: this.getNameConflictBehaviorValue(args.options.nameConflictBehavior),
                newName: args.options.newName,
                operation: 'move'
            });
            if (args.options.skipWait) {
                return;
            }
            if (this.verbose) {
                await logger.logToStderr('Waiting for the move job to complete...');
            }
            const copyJobResult = await spo.getCopyJobResult(args.options.webUrl, copyJobResponse);
            if (this.verbose) {
                await logger.logToStderr('Getting information about the destination folder...');
            }
            // Get destination folder data
            const siteRelativeDestinationFolder = '/' + copyJobResult.TargetObjectSiteRelativeUrl.substring(0, copyJobResult.TargetObjectSiteRelativeUrl.lastIndexOf('/'));
            const absoluteWebUrl = destinationPath.substring(0, destinationPath.toLowerCase().lastIndexOf(siteRelativeDestinationFolder.toLowerCase()));
            const requestOptions = {
                url: `${absoluteWebUrl}/_api/Web/GetFolderById('${copyJobResult.TargetObjectUniqueId}')`,
                headers: {
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            const destinationFile = await request.get(requestOptions);
            await logger.log(destinationFile);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getNameConflictBehaviorValue(nameConflictBehavior) {
        switch (nameConflictBehavior?.toLowerCase()) {
            case 'fail':
                return CreateFolderCopyJobsNameConflictBehavior.Fail;
            case 'rename':
                return CreateFolderCopyJobsNameConflictBehavior.Rename;
            default:
                return CreateFolderCopyJobsNameConflictBehavior.Fail;
        }
    }
    async getSourcePath(logger, options) {
        if (options.sourceUrl) {
            return urlUtil.getServerRelativePath(options.webUrl, options.sourceUrl);
        }
        if (this.verbose) {
            await logger.logToStderr(`Retrieving server-relative path for folder with ID '${options.sourceId}'...`);
        }
        const requestOptions = {
            url: `${options.webUrl}/_api/Web/GetFolderById('${options.sourceId}')/ServerRelativePath`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const path = await request.get(requestOptions);
        return path.DecodedUrl;
    }
    getAbsoluteUrl(webUrl, url) {
        const result = url.startsWith('https://') ? url : urlUtil.getAbsoluteUrl(webUrl, url);
        return urlUtil.removeTrailingSlashes(result);
    }
}
_SpoFolderMoveCommand_instances = new WeakSet(), _SpoFolderMoveCommand_initTelemetry = function _SpoFolderMoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            sourceUrl: typeof args.options.sourceUrl !== 'undefined',
            sourceId: typeof args.options.sourceId !== 'undefined',
            newName: typeof args.options.newName !== 'undefined',
            nameConflictBehavior: typeof args.options.nameConflictBehavior !== 'undefined',
            skipWait: !!args.options.skipWait
        });
    });
}, _SpoFolderMoveCommand_initOptions = function _SpoFolderMoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-s, --sourceUrl [sourceUrl]'
    }, {
        option: '-i, --sourceId [sourceId]'
    }, {
        option: '-t, --targetUrl <targetUrl>'
    }, {
        option: '--newName [newName]'
    }, {
        option: '--nameConflictBehavior [nameConflictBehavior]',
        autocomplete: this.nameConflictBehaviorOptions
    }, {
        option: '--skipWait'
    });
}, _SpoFolderMoveCommand_initValidators = function _SpoFolderMoveCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.sourceId && !validation.isValidGuid(args.options.sourceId)) {
            return `'${args.options.sourceId}' is not a valid GUID for sourceId.`;
        }
        if (args.options.nameConflictBehavior && this.nameConflictBehaviorOptions.indexOf(args.options.nameConflictBehavior) === -1) {
            return `'${args.options.nameConflictBehavior}' is not a valid value for nameConflictBehavior. Allowed values are: ${this.nameConflictBehaviorOptions.join(', ')}.`;
        }
        return true;
    });
}, _SpoFolderMoveCommand_initOptionSets = function _SpoFolderMoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['sourceUrl', 'sourceId'] });
}, _SpoFolderMoveCommand_initTypes = function _SpoFolderMoveCommand_initTypes() {
    this.types.string.push('webUrl', 'sourceUrl', 'sourceId', 'targetUrl', 'newName', 'nameConflictBehavior');
    this.types.boolean.push('skipWait');
};
export default new SpoFolderMoveCommand();
//# sourceMappingURL=folder-move.js.map