var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFileRemoveCommand_instances, _SpoFileRemoveCommand_initTelemetry, _SpoFileRemoveCommand_initOptions, _SpoFileRemoveCommand_initValidators, _SpoFileRemoveCommand_initOptionSets, _SpoFileRemoveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFileRemoveCommand extends SpoCommand {
    get name() {
        return commands.FILE_REMOVE;
    }
    get description() {
        return 'Removes the specified file';
    }
    alias() {
        return [commands.PAGE_TEMPLATE_REMOVE];
    }
    constructor() {
        super();
        _SpoFileRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFileRemoveCommand_instances, "m", _SpoFileRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFileRemoveCommand_instances, "m", _SpoFileRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFileRemoveCommand_instances, "m", _SpoFileRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFileRemoveCommand_instances, "m", _SpoFileRemoveCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFileRemoveCommand_instances, "m", _SpoFileRemoveCommand_initTypes).call(this);
    }
    getExcludedOptionsWithUrls() {
        return ['url'];
    }
    async commandAction(logger, args) {
        const removeFile = async () => {
            if (this.verbose) {
                await logger.logToStderr(`Removing file in site at ${args.options.webUrl}...`);
            }
            let requestUrl;
            if (args.options.id) {
                requestUrl = `${args.options.webUrl}/_api/web/GetFileById(guid'${formatting.encodeQueryParameter(args.options.id)}')`;
            }
            else {
                const serverRelativePath = urlUtil.getServerRelativePath(args.options.webUrl, args.options.url);
                requestUrl = `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')`;
            }
            if (args.options.recycle) {
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
            if (args.options.bypassSharedLock) {
                requestOptions.headers.Prefer = 'bypass-shared-lock';
            }
            try {
                await request.post(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeFile();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to ${args.options.recycle ? 'recycle' : 'remove'} the file ${args.options.id || args.options.url} located in site ${args.options.webUrl}?` });
            if (result) {
                await removeFile();
            }
        }
    }
}
_SpoFileRemoveCommand_instances = new WeakSet(), _SpoFileRemoveCommand_initTelemetry = function _SpoFileRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            url: typeof args.options.url !== 'undefined',
            recycle: !!args.options.recycle,
            bypassSharedLock: !!args.options.bypassSharedLock,
            force: !!args.options.force
        });
    });
}, _SpoFileRemoveCommand_initOptions = function _SpoFileRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '--url [url]'
    }, {
        option: '--recycle'
    }, {
        option: '--bypassSharedLock'
    }, {
        option: '-f, --force'
    });
}, _SpoFileRemoveCommand_initValidators = function _SpoFileRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.id &&
            !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
}, _SpoFileRemoveCommand_initOptionSets = function _SpoFileRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'url'] });
}, _SpoFileRemoveCommand_initTypes = function _SpoFileRemoveCommand_initTypes() {
    this.types.string.push('webUrl', 'id', 'url');
    this.types.boolean.push('recycle', 'bypassSharedLock', 'force');
};
export default new SpoFileRemoveCommand();
//# sourceMappingURL=file-remove.js.map