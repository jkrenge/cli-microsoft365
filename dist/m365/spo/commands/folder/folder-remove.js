var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFolderRemoveCommand_instances, _SpoFolderRemoveCommand_initTelemetry, _SpoFolderRemoveCommand_initOptions, _SpoFolderRemoveCommand_initValidators, _SpoFolderRemoveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFolderRemoveCommand extends SpoCommand {
    get name() {
        return commands.FOLDER_REMOVE;
    }
    get description() {
        return 'Deletes the specified folder';
    }
    constructor() {
        super();
        _SpoFolderRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFolderRemoveCommand_instances, "m", _SpoFolderRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFolderRemoveCommand_instances, "m", _SpoFolderRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFolderRemoveCommand_instances, "m", _SpoFolderRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFolderRemoveCommand_instances, "m", _SpoFolderRemoveCommand_initTypes).call(this);
    }
    getExcludedOptionsWithUrls() {
        return ['url'];
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeFolder(logger, args.options);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to ${args.options.recycle ? "recycle" : "remove"} the folder ${args.options.url} located in site ${args.options.webUrl}?` });
            if (result) {
                await this.removeFolder(logger, args.options);
            }
        }
    }
    async removeFolder(logger, options) {
        if (this.verbose) {
            await logger.logToStderr(`Removing folder in site at ${options.webUrl}...`);
        }
        const serverRelativePath = urlUtil.getServerRelativePath(options.webUrl, options.url);
        let requestUrl = `${options.webUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')`;
        if (options.recycle) {
            requestUrl += `/recycle()`;
        }
        const requestOptions = {
            url: requestUrl,
            method: 'POST',
            headers: {
                'X-HTTP-Method': 'DELETE',
                'If-Match': '*',
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        try {
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoFolderRemoveCommand_instances = new WeakSet(), _SpoFolderRemoveCommand_initTelemetry = function _SpoFolderRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            recycle: !!args.options.recycle,
            force: !!args.options.force
        });
    });
}, _SpoFolderRemoveCommand_initOptions = function _SpoFolderRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--url <url>'
    }, {
        option: '--recycle'
    }, {
        option: '-f, --force'
    });
}, _SpoFolderRemoveCommand_initValidators = function _SpoFolderRemoveCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.webUrl));
}, _SpoFolderRemoveCommand_initTypes = function _SpoFolderRemoveCommand_initTypes() {
    this.types.string.push('webUrl', 'url');
    this.types.boolean.push('recycle', 'force');
};
export default new SpoFolderRemoveCommand();
//# sourceMappingURL=folder-remove.js.map