var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFileVersionListCommand_instances, _SpoFileVersionListCommand_initTelemetry, _SpoFileVersionListCommand_initOptions, _SpoFileVersionListCommand_initValidators, _SpoFileVersionListCommand_initOptionSets, _SpoFileVersionListCommand_initTypes;
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFileVersionListCommand extends SpoCommand {
    get name() {
        return commands.FILE_VERSION_LIST;
    }
    get description() {
        return 'Retrieves all versions of a file';
    }
    defaultProperties() {
        return ['Created', 'ID', 'IsCurrentVersion', 'VersionLabel', 'ExpirationDate'];
    }
    constructor() {
        super();
        _SpoFileVersionListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFileVersionListCommand_instances, "m", _SpoFileVersionListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFileVersionListCommand_instances, "m", _SpoFileVersionListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFileVersionListCommand_instances, "m", _SpoFileVersionListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFileVersionListCommand_instances, "m", _SpoFileVersionListCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFileVersionListCommand_instances, "m", _SpoFileVersionListCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving all versions of file ${args.options.fileUrl || args.options.fileId} at site ${args.options.webUrl}...`);
        }
        try {
            let requestUrl = `${args.options.webUrl}/_api/web`;
            if (args.options.fileUrl) {
                const serverRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl);
                requestUrl += `/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}')`;
            }
            else {
                requestUrl += `/GetFileById('${args.options.fileId}')`;
            }
            requestUrl += `/versions?$select=*,ExpirationDate`;
            const response = await odata.getAllItems(requestUrl);
            await logger.log(response);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoFileVersionListCommand_instances = new WeakSet(), _SpoFileVersionListCommand_initTelemetry = function _SpoFileVersionListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            fileUrl: typeof args.options.fileUrl !== 'undefined',
            fileId: typeof args.options.fileId !== 'undefined'
        });
    });
}, _SpoFileVersionListCommand_initOptions = function _SpoFileVersionListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--fileUrl [fileUrl]'
    }, {
        option: '-i, --fileId [fileId]'
    });
}, _SpoFileVersionListCommand_initValidators = function _SpoFileVersionListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.fileId && !validation.isValidGuid(args.options.fileId)) {
            return `${args.options.fileId} is not a valid GUID`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
}, _SpoFileVersionListCommand_initOptionSets = function _SpoFileVersionListCommand_initOptionSets() {
    this.optionSets.push({ options: ['fileUrl', 'fileId'] });
}, _SpoFileVersionListCommand_initTypes = function _SpoFileVersionListCommand_initTypes() {
    this.types.string.push('webUrl', 'fileUrl', 'fileId');
};
export default new SpoFileVersionListCommand();
//# sourceMappingURL=file-version-list.js.map