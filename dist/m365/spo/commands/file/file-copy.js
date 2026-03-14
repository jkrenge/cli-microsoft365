var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFileCopyCommand_instances, _SpoFileCopyCommand_initTelemetry, _SpoFileCopyCommand_initOptions, _SpoFileCopyCommand_initValidators, _SpoFileCopyCommand_initOptionSets, _SpoFileCopyCommand_initTypes;
import request from '../../../../request.js';
import { CreateFileCopyJobsNameConflictBehavior, spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFileCopyCommand extends SpoCommand {
    get name() {
        return commands.FILE_COPY;
    }
    get description() {
        return 'Copies a file to another location';
    }
    constructor() {
        super();
        _SpoFileCopyCommand_instances.add(this);
        this.nameConflictBehaviorOptions = ['fail', 'replace', 'rename'];
        __classPrivateFieldGet(this, _SpoFileCopyCommand_instances, "m", _SpoFileCopyCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFileCopyCommand_instances, "m", _SpoFileCopyCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFileCopyCommand_instances, "m", _SpoFileCopyCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFileCopyCommand_instances, "m", _SpoFileCopyCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFileCopyCommand_instances, "m", _SpoFileCopyCommand_initTypes).call(this);
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
                await logger.logToStderr(`Copying file '${sourceServerRelativePath}' to '${args.options.targetUrl}'...`);
            }
            let newName = args.options.newName;
            // Add original file extension if not provided
            if (newName && !newName.includes('.')) {
                newName += sourceServerRelativePath.substring(sourceServerRelativePath.lastIndexOf('.'));
            }
            const copyJobResponse = await spo.createFileCopyJob(args.options.webUrl, sourcePath, destinationPath, {
                nameConflictBehavior: this.getNameConflictBehaviorValue(args.options.nameConflictBehavior),
                bypassSharedLock: !!args.options.bypassSharedLock,
                ignoreVersionHistory: !!args.options.ignoreVersionHistory,
                newName: newName,
                operation: 'copy'
            });
            if (args.options.skipWait) {
                return;
            }
            if (this.verbose) {
                await logger.logToStderr('Waiting for the copy job to complete...');
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
_SpoFileCopyCommand_instances = new WeakSet(), _SpoFileCopyCommand_initTelemetry = function _SpoFileCopyCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            sourceUrl: typeof args.options.sourceUrl !== 'undefined',
            sourceId: typeof args.options.sourceId !== 'undefined',
            newName: typeof args.options.newName !== 'undefined',
            nameConflictBehavior: typeof args.options.nameConflictBehavior !== 'undefined',
            ignoreVersionHistory: !!args.options.ignoreVersionHistory,
            bypassSharedLock: !!args.options.bypassSharedLock,
            skipWait: !!args.options.skipWait
        });
    });
}, _SpoFileCopyCommand_initOptions = function _SpoFileCopyCommand_initOptions() {
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
        option: '--bypassSharedLock'
    }, {
        option: '--ignoreVersionHistory'
    }, {
        option: '--skipWait'
    });
}, _SpoFileCopyCommand_initValidators = function _SpoFileCopyCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.sourceId && !validation.isValidGuid(args.options.sourceId)) {
            return `${args.options.sourceId} is not a valid GUID for sourceId.`;
        }
        if (args.options.nameConflictBehavior && !this.nameConflictBehaviorOptions.includes(args.options.nameConflictBehavior.toLowerCase())) {
            return `${args.options.nameConflictBehavior} is not a valid nameConflictBehavior value. Allowed values: ${this.nameConflictBehaviorOptions.join(', ')}.`;
        }
        return true;
    });
}, _SpoFileCopyCommand_initOptionSets = function _SpoFileCopyCommand_initOptionSets() {
    this.optionSets.push({ options: ['sourceUrl', 'sourceId'] });
}, _SpoFileCopyCommand_initTypes = function _SpoFileCopyCommand_initTypes() {
    this.types.string.push('webUrl', 'sourceUrl', 'sourceId', 'targetUrl', 'newName', 'nameConflictBehavior');
    this.types.boolean.push('bypassSharedLock', 'ignoreVersionHistory', 'skipWait');
};
export default new SpoFileCopyCommand();
//# sourceMappingURL=file-copy.js.map