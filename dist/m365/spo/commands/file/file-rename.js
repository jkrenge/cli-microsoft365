var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFileRenameCommand_instances, _SpoFileRenameCommand_initTelemetry, _SpoFileRenameCommand_initOptions, _SpoFileRenameCommand_initValidators, _SpoFileRenameCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import removeCommand from './file-remove.js';
class SpoFileRenameCommand extends SpoCommand {
    get name() {
        return commands.FILE_RENAME;
    }
    get description() {
        return 'Renames a file';
    }
    constructor() {
        super();
        _SpoFileRenameCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFileRenameCommand_instances, "m", _SpoFileRenameCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFileRenameCommand_instances, "m", _SpoFileRenameCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFileRenameCommand_instances, "m", _SpoFileRenameCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFileRenameCommand_instances, "m", _SpoFileRenameCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const webUrl = args.options.webUrl;
        const originalFileServerRelativePath = urlUtil.getServerRelativePath(args.options.webUrl, args.options.sourceUrl);
        try {
            await this.getFile(originalFileServerRelativePath, webUrl);
            if (args.options.force) {
                await this.deleteFile(webUrl, args.options.sourceUrl, args.options.targetFileName);
            }
            const requestBody = {
                formValues: [{
                        FieldName: 'FileLeafRef',
                        FieldValue: args.options.targetFileName
                    }]
            };
            const requestOptions = {
                url: `${webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(originalFileServerRelativePath)}')/ListItemAllFields/ValidateUpdateListItem()`,
                headers: {
                    'accept': 'application/json;odata=nometadata'
                },
                data: requestBody,
                responseType: 'json'
            };
            const resp = await request.post(requestOptions);
            await logger.log(resp.value);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getFile(originalFileServerRelativeUrl, webUrl) {
        const requestUrl = `${webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(originalFileServerRelativeUrl)}')?$select=UniqueId`;
        const requestOptions = {
            url: requestUrl,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        return request.get(requestOptions);
    }
    async deleteFile(webUrl, sourceUrl, targetFileName) {
        const targetFileServerRelativeUrl = `${urlUtil.getServerRelativePath(webUrl, sourceUrl.substring(0, sourceUrl.lastIndexOf('/')))}/${targetFileName}`;
        const removeOptions = {
            webUrl: webUrl,
            url: targetFileServerRelativeUrl,
            recycle: true,
            force: true,
            debug: this.debug,
            verbose: this.verbose
        };
        try {
            await cli.executeCommand(removeCommand, { options: { ...removeOptions, _: [] } });
        }
        catch (err) {
            if (err?.error?.message?.includes('does not exist')) {
                return;
            }
            throw err;
        }
    }
}
_SpoFileRenameCommand_instances = new WeakSet(), _SpoFileRenameCommand_initTelemetry = function _SpoFileRenameCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: !!args.options.force
        });
    });
}, _SpoFileRenameCommand_initOptions = function _SpoFileRenameCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-s, --sourceUrl <sourceUrl>'
    }, {
        option: '-t, --targetFileName <targetFileName>'
    }, {
        option: '--force'
    });
}, _SpoFileRenameCommand_initValidators = function _SpoFileRenameCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.webUrl));
}, _SpoFileRenameCommand_initTypes = function _SpoFileRenameCommand_initTypes() {
    this.types.string.push('webUrl', 'sourceUrl', 'targetFileName');
    this.types.boolean.push('force');
};
export default new SpoFileRenameCommand();
//# sourceMappingURL=file-rename.js.map