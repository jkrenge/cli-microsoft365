var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFileCheckoutUndoCommand_instances, _SpoFileCheckoutUndoCommand_initTelemetry, _SpoFileCheckoutUndoCommand_initOptions, _SpoFileCheckoutUndoCommand_initValidators, _SpoFileCheckoutUndoCommand_initOptionSets, _SpoFileCheckoutUndoCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFileCheckoutUndoCommand extends SpoCommand {
    get name() {
        return commands.FILE_CHECKOUT_UNDO;
    }
    get description() {
        return 'Discards a checked out file';
    }
    constructor() {
        super();
        _SpoFileCheckoutUndoCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFileCheckoutUndoCommand_instances, "m", _SpoFileCheckoutUndoCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFileCheckoutUndoCommand_instances, "m", _SpoFileCheckoutUndoCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFileCheckoutUndoCommand_instances, "m", _SpoFileCheckoutUndoCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFileCheckoutUndoCommand_instances, "m", _SpoFileCheckoutUndoCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFileCheckoutUndoCommand_instances, "m", _SpoFileCheckoutUndoCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const undoCheckout = async () => {
            try {
                if (this.verbose) {
                    await logger.logToStderr(`Undoing checkout for file ${args.options.fileId || args.options.fileUrl} on web ${args.options.webUrl}`);
                }
                let requestUrl = `${args.options.webUrl}/_api/web/`;
                if (args.options.fileId) {
                    requestUrl += `getFileById('${args.options.fileId}')`;
                }
                else if (args.options.fileUrl) {
                    const serverRelativePath = urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl);
                    requestUrl += `getFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')`;
                }
                requestUrl += '/undocheckout';
                const requestOptions = {
                    url: requestUrl,
                    headers: {
                        'accept': 'application/json;odata=nometadata'
                    },
                    responseType: 'json'
                };
                await request.post(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await undoCheckout();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to undo the checkout for file ${args.options.fileId || args.options.fileUrl}?` });
            if (result) {
                await undoCheckout();
            }
        }
    }
}
_SpoFileCheckoutUndoCommand_instances = new WeakSet(), _SpoFileCheckoutUndoCommand_initTelemetry = function _SpoFileCheckoutUndoCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            fileId: typeof args.options.fileId !== 'undefined',
            fileUrl: typeof args.options.fileUrl !== 'undefined',
            force: !!args.options.force
        });
    });
}, _SpoFileCheckoutUndoCommand_initOptions = function _SpoFileCheckoutUndoCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--fileUrl [fileUrl]'
    }, {
        option: '-i, --fileId [fileId]'
    }, {
        option: '-f, --force'
    });
}, _SpoFileCheckoutUndoCommand_initValidators = function _SpoFileCheckoutUndoCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.fileId && !validation.isValidGuid(args.options.fileId)) {
            return `${args.options.fileId} is not a valid GUID`;
        }
        return true;
    });
}, _SpoFileCheckoutUndoCommand_initOptionSets = function _SpoFileCheckoutUndoCommand_initOptionSets() {
    this.optionSets.push({ options: ['fileId', 'fileUrl'] });
}, _SpoFileCheckoutUndoCommand_initTypes = function _SpoFileCheckoutUndoCommand_initTypes() {
    this.types.string.push('webUrl', 'fileUrl', 'fileId');
    this.types.boolean.push('force');
};
export default new SpoFileCheckoutUndoCommand();
//# sourceMappingURL=file-checkout-undo.js.map