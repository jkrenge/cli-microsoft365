var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFileVersionGetCommand_instances, _SpoFileVersionGetCommand_initTelemetry, _SpoFileVersionGetCommand_initOptions, _SpoFileVersionGetCommand_initValidators, _SpoFileVersionGetCommand_initOptionSets, _SpoFileVersionGetCommand_initTypes;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFileVersionGetCommand extends SpoCommand {
    get name() {
        return commands.FILE_VERSION_GET;
    }
    get description() {
        return 'Get a specific version of a file in a SharePoint Document library';
    }
    constructor() {
        super();
        _SpoFileVersionGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFileVersionGetCommand_instances, "m", _SpoFileVersionGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFileVersionGetCommand_instances, "m", _SpoFileVersionGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFileVersionGetCommand_instances, "m", _SpoFileVersionGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFileVersionGetCommand_instances, "m", _SpoFileVersionGetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFileVersionGetCommand_instances, "m", _SpoFileVersionGetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving version ${args.options.label} of the file ${args.options.fileUrl || args.options.fileId} at site ${args.options.webUrl}...`);
        }
        try {
            const version = await this.getVersion(args);
            await logger.log(version.value[0]);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getVersion(args) {
        let requestUrl = `${args.options.webUrl}/_api/web/`;
        if (args.options.fileUrl) {
            const serverRelUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl);
            requestUrl += `GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelUrl)}')/versions/?$filter=VersionLabel eq '${args.options.label}'`;
        }
        else {
            requestUrl += `GetFileById('${args.options.fileId}')/versions/?$filter=VersionLabel eq '${args.options.label}'`;
        }
        requestUrl += `&$select=*,ExpirationDate`;
        const requestOptions = {
            url: requestUrl,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        return response;
    }
}
_SpoFileVersionGetCommand_instances = new WeakSet(), _SpoFileVersionGetCommand_initTelemetry = function _SpoFileVersionGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            fileUrl: args.options.fileUrl,
            fileId: args.options.fileId
        });
    });
}, _SpoFileVersionGetCommand_initOptions = function _SpoFileVersionGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--label <label>'
    }, {
        option: '--fileUrl [fileUrl]'
    }, {
        option: '-i, --fileId [fileId]'
    });
}, _SpoFileVersionGetCommand_initValidators = function _SpoFileVersionGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.fileId && !validation.isValidGuid(args.options.fileId)) {
            return `${args.options.fileId} is not a valid GUID`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
}, _SpoFileVersionGetCommand_initOptionSets = function _SpoFileVersionGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['fileUrl', 'fileId'] });
}, _SpoFileVersionGetCommand_initTypes = function _SpoFileVersionGetCommand_initTypes() {
    this.types.string.push('webUrl', 'label', 'fileUrl', 'fileId');
};
export default new SpoFileVersionGetCommand();
//# sourceMappingURL=file-version-get.js.map