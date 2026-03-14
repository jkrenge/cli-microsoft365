var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFileVersionRemoveCommand_instances, _SpoFileVersionRemoveCommand_initTelemetry, _SpoFileVersionRemoveCommand_initOptions, _SpoFileVersionRemoveCommand_initValidators, _SpoFileVersionRemoveCommand_initOptionSets, _SpoFileVersionRemoveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFileVersionRemoveCommand extends SpoCommand {
    get name() {
        return commands.FILE_VERSION_REMOVE;
    }
    get description() {
        return 'Removes a specific version of a file in a SharePoint Document library';
    }
    constructor() {
        super();
        _SpoFileVersionRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFileVersionRemoveCommand_instances, "m", _SpoFileVersionRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFileVersionRemoveCommand_instances, "m", _SpoFileVersionRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFileVersionRemoveCommand_instances, "m", _SpoFileVersionRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFileVersionRemoveCommand_instances, "m", _SpoFileVersionRemoveCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFileVersionRemoveCommand_instances, "m", _SpoFileVersionRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Removes version ${args.options.label} of the file ${args.options.fileUrl || args.options.fileId} at site ${args.options.webUrl}...`);
        }
        try {
            if (args.options.force) {
                await this.removeVersion(args);
            }
            else {
                const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the version ${args.options.label} from file ${args.options.fileId || args.options.fileUrl}'?` });
                if (result) {
                    await this.removeVersion(args);
                }
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async removeVersion(args) {
        let requestUrl = `${args.options.webUrl}/_api/web/`;
        if (args.options.fileUrl) {
            const serverRelativePath = urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl);
            requestUrl += `GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')/versions/DeleteByLabel('${args.options.label}')`;
        }
        else {
            requestUrl += `GetFileById('${args.options.fileId}')/versions/DeleteByLabel('${args.options.label}')`;
        }
        const requestOptions = {
            url: requestUrl,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        return request.delete(requestOptions);
    }
}
_SpoFileVersionRemoveCommand_instances = new WeakSet(), _SpoFileVersionRemoveCommand_initTelemetry = function _SpoFileVersionRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            fileUrl: args.options.fileUrl,
            fileId: args.options.fileId,
            force: (!!args.options.force).toString()
        });
    });
}, _SpoFileVersionRemoveCommand_initOptions = function _SpoFileVersionRemoveCommand_initOptions() {
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
}, _SpoFileVersionRemoveCommand_initValidators = function _SpoFileVersionRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.fileId && !validation.isValidGuid(args.options.fileId)) {
            return `${args.options.fileId} is not a valid GUID`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
}, _SpoFileVersionRemoveCommand_initOptionSets = function _SpoFileVersionRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['fileUrl', 'fileId'] });
}, _SpoFileVersionRemoveCommand_initTypes = function _SpoFileVersionRemoveCommand_initTypes() {
    this.types.string.push('webUrl', 'label', 'fileUrl', 'fileId');
    this.types.boolean.push('force');
};
export default new SpoFileVersionRemoveCommand();
//# sourceMappingURL=file-version-remove.js.map