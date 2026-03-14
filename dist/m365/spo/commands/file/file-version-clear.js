var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFileVersionClearCommand_instances, _SpoFileVersionClearCommand_initTelemetry, _SpoFileVersionClearCommand_initOptions, _SpoFileVersionClearCommand_initValidators, _SpoFileVersionClearCommand_initOptionSets, _SpoFileVersionClearCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFileVersionClearCommand extends SpoCommand {
    get name() {
        return commands.FILE_VERSION_CLEAR;
    }
    get description() {
        return 'Deletes all file version history of a file in a SharePoint Document library';
    }
    constructor() {
        super();
        _SpoFileVersionClearCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFileVersionClearCommand_instances, "m", _SpoFileVersionClearCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFileVersionClearCommand_instances, "m", _SpoFileVersionClearCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFileVersionClearCommand_instances, "m", _SpoFileVersionClearCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFileVersionClearCommand_instances, "m", _SpoFileVersionClearCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFileVersionClearCommand_instances, "m", _SpoFileVersionClearCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Deletes all version history of the file ${args.options.fileUrl || args.options.fileId} at site ${args.options.webUrl}...`);
        }
        try {
            if (args.options.force) {
                await this.clearVersions(args);
            }
            else {
                const result = await cli.promptForConfirmation({ message: `Are you sure you want to delete all version history for file ${args.options.fileId || args.options.fileUrl}'?` });
                if (result) {
                    await this.clearVersions(args);
                }
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async clearVersions(args) {
        let requestUrl = `${args.options.webUrl}/_api/web/`;
        if (args.options.fileUrl) {
            const serverRelativePath = urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl);
            requestUrl += `GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')/versions/DeleteAll()`;
        }
        else {
            requestUrl += `GetFileById('${args.options.fileId}')/versions/DeleteAll()`;
        }
        const requestOptions = {
            url: requestUrl,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        return request.post(requestOptions);
    }
}
_SpoFileVersionClearCommand_instances = new WeakSet(), _SpoFileVersionClearCommand_initTelemetry = function _SpoFileVersionClearCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            fileUrl: args.options.fileUrl,
            fileId: args.options.fileId,
            force: !!args.options.force
        });
    });
}, _SpoFileVersionClearCommand_initOptions = function _SpoFileVersionClearCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--fileUrl [fileUrl]'
    }, {
        option: '-i, --fileId [fileId]'
    }, {
        option: '-f, --force'
    });
}, _SpoFileVersionClearCommand_initValidators = function _SpoFileVersionClearCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.fileId && !validation.isValidGuid(args.options.fileId)) {
            return `${args.options.fileId} is not a valid GUID`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
}, _SpoFileVersionClearCommand_initOptionSets = function _SpoFileVersionClearCommand_initOptionSets() {
    this.optionSets.push({ options: ['fileUrl', 'fileId'] });
}, _SpoFileVersionClearCommand_initTypes = function _SpoFileVersionClearCommand_initTypes() {
    this.types.string.push('webUrl', 'fileUrl', 'fileId');
    this.types.boolean.push('force');
};
export default new SpoFileVersionClearCommand();
//# sourceMappingURL=file-version-clear.js.map