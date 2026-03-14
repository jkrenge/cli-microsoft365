var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFileVersionRestoreCommand_instances, _SpoFileVersionRestoreCommand_initTelemetry, _SpoFileVersionRestoreCommand_initOptions, _SpoFileVersionRestoreCommand_initValidators, _SpoFileVersionRestoreCommand_initOptionSets, _SpoFileVersionRestoreCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFileVersionRestoreCommand extends SpoCommand {
    get name() {
        return commands.FILE_VERSION_RESTORE;
    }
    get description() {
        return 'Restores a specific version of a file in a SharePoint Document library';
    }
    constructor() {
        super();
        _SpoFileVersionRestoreCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFileVersionRestoreCommand_instances, "m", _SpoFileVersionRestoreCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFileVersionRestoreCommand_instances, "m", _SpoFileVersionRestoreCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFileVersionRestoreCommand_instances, "m", _SpoFileVersionRestoreCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFileVersionRestoreCommand_instances, "m", _SpoFileVersionRestoreCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFileVersionRestoreCommand_instances, "m", _SpoFileVersionRestoreCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Restores version ${args.options.label} of the file ${args.options.fileUrl || args.options.fileId} at site ${args.options.webUrl}...`);
        }
        try {
            if (args.options.force) {
                await this.restoreVersion(args);
            }
            else {
                const result = await cli.promptForConfirmation({ message: `Are you sure you want to restore the version ${args.options.label} from file ${args.options.fileId || args.options.fileUrl}'?` });
                if (result) {
                    await this.restoreVersion(args);
                }
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async restoreVersion(args) {
        let requestUrl = `${args.options.webUrl}/_api/web/`;
        if (args.options.fileUrl) {
            const serverRelativePath = urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl);
            requestUrl += `GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')/versions/RestoreByLabel('${args.options.label}')`;
        }
        else {
            requestUrl += `GetFileById('${args.options.fileId}')/versions/RestoreByLabel('${args.options.label}')`;
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
_SpoFileVersionRestoreCommand_instances = new WeakSet(), _SpoFileVersionRestoreCommand_initTelemetry = function _SpoFileVersionRestoreCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            fileUrl: args.options.fileUrl,
            fileId: args.options.fileId,
            force: (!!args.options.force).toString()
        });
    });
}, _SpoFileVersionRestoreCommand_initOptions = function _SpoFileVersionRestoreCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--label <label>'
    }, {
        option: '--fileUrl [fileUrl]'
    }, {
        option: '-i, --fileId [fileId]'
    }, {
        option: '-f, --force'
    });
}, _SpoFileVersionRestoreCommand_initValidators = function _SpoFileVersionRestoreCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.fileId && !validation.isValidGuid(args.options.fileId)) {
            return `${args.options.fileId} is not a valid GUID`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
}, _SpoFileVersionRestoreCommand_initOptionSets = function _SpoFileVersionRestoreCommand_initOptionSets() {
    this.optionSets.push({ options: ['fileUrl', 'fileId'] });
}, _SpoFileVersionRestoreCommand_initTypes = function _SpoFileVersionRestoreCommand_initTypes() {
    this.types.string.push('webUrl', 'label', 'fileUrl', 'fileId');
    this.types.boolean.push('force');
};
export default new SpoFileVersionRestoreCommand();
//# sourceMappingURL=file-version-restore.js.map