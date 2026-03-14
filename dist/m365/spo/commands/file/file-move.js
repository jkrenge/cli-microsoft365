var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFileMoveCommand_instances, _SpoFileMoveCommand_initTelemetry, _SpoFileMoveCommand_initOptions, _SpoFileMoveCommand_initValidators, _SpoFileMoveCommand_initOptionSets, _SpoFileMoveCommand_initTypes;
import request from '../../../../request.js';
import { CreateFileCopyJobsNameConflictBehavior, spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFileMoveCommand extends SpoCommand {
    get name() {
        return commands.FILE_MOVE;
    }
    get description() {
        return 'Moves a file to another location';
    }
    constructor() {
        super();
        _SpoFileMoveCommand_instances.add(this);
        this.nameConflictBehaviorOptions = ['fail', 'replace', 'rename'];
        __classPrivateFieldGet(this, _SpoFileMoveCommand_instances, "m", _SpoFileMoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFileMoveCommand_instances, "m", _SpoFileMoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFileMoveCommand_instances, "m", _SpoFileMoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFileMoveCommand_instances, "m", _SpoFileMoveCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFileMoveCommand_instances, "m", _SpoFileMoveCommand_initTypes).call(this);
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
                await logger.logToStderr(`Moving file '${sourceServerRelativePath}' to '${args.options.targetUrl}'...`);
            }
            let newName = args.options.newName;
            // Add original file extension if not provided
            if (newName && !newName.includes('.')) {
                newName += sourceServerRelativePath.substring(sourceServerRelativePath.lastIndexOf('.'));
            }
            const copyJobResponse = await spo.createFileCopyJob(args.options.webUrl, sourcePath, destinationPath, {
                nameConflictBehavior: this.getNameConflictBehaviorValue(args.options.nameConflictBehavior),
                bypassSharedLock: !!args.options.bypassSharedLock,
                includeItemPermissions: !!args.options.withItemPermissions,
                newName: newName,
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
                await logger.logToStderr('Getting information about the destination file...');
            }
            // Get destination file data
            const siteRelativeDestinationFolder = '/' + copyJobResult.TargetObjectSiteRelativeUrl.substring(0, copyJobResult.TargetObjectSiteRelativeUrl.lastIndexOf('/'));
            const absoluteWebUrl = destinationPath.substring(0, destinationPath.toLowerCase().lastIndexOf(siteRelativeDestinationFolder.toLowerCase()));
            const requestOptions = {
                url: `${absoluteWebUrl}/_api/Web/GetFileById('${copyJobResult.TargetObjectUniqueId}')`,
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
    async getSourcePath(logger, options) {
        if (options.sourceUrl) {
            return urlUtil.getServerRelativePath(options.webUrl, options.sourceUrl);
        }
        if (this.verbose) {
            await logger.logToStderr(`Retrieving server-relative path for file with ID '${options.sourceId}'...`);
        }
        const requestOptions = {
            url: `${options.webUrl}/_api/Web/GetFileById('${options.sourceId}')/ServerRelativePath`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const file = await request.get(requestOptions);
        return file.DecodedUrl;
    }
    getNameConflictBehaviorValue(nameConflictBehavior) {
        switch (nameConflictBehavior?.toLowerCase()) {
            case 'fail':
                return CreateFileCopyJobsNameConflictBehavior.Fail;
            case 'replace':
                return CreateFileCopyJobsNameConflictBehavior.Replace;
            case 'rename':
                return CreateFileCopyJobsNameConflictBehavior.Rename;
            default:
                return CreateFileCopyJobsNameConflictBehavior.Fail;
        }
    }
    getAbsoluteUrl(webUrl, url) {
        const result = url.startsWith('https://') ? url : urlUtil.getAbsoluteUrl(webUrl, url);
        return urlUtil.removeTrailingSlashes(result);
    }
}
_SpoFileMoveCommand_instances = new WeakSet(), _SpoFileMoveCommand_initTelemetry = function _SpoFileMoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            sourceUrl: typeof args.options.sourceUrl !== 'undefined',
            sourceId: typeof args.options.sourceId !== 'undefined',
            newName: typeof args.options.newName !== 'undefined',
            nameConflictBehavior: typeof args.options.nameConflictBehavior !== 'undefined',
            withItemPermissions: !!args.options.withItemPermissions,
            bypassSharedLock: !!args.options.bypassSharedLock,
            skipWait: !!args.options.skipWait
        });
    });
}, _SpoFileMoveCommand_initOptions = function _SpoFileMoveCommand_initOptions() {
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
        option: '--withItemPermissions'
    }, {
        option: '--bypassSharedLock'
    }, {
        option: '--skipWait'
    });
}, _SpoFileMoveCommand_initValidators = function _SpoFileMoveCommand_initValidators() {
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
}, _SpoFileMoveCommand_initOptionSets = function _SpoFileMoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['sourceUrl', 'sourceId'] });
}, _SpoFileMoveCommand_initTypes = function _SpoFileMoveCommand_initTypes() {
    this.types.string.push('webUrl', 'sourceUrl', 'sourceId', 'targetUrl', 'newName', 'nameConflictBehavior');
    this.types.boolean.push('withItemPermissions', 'bypassSharedLock', 'skipWait');
};
export default new SpoFileMoveCommand();
//# sourceMappingURL=file-move.js.map