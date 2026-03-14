var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFileRoleInheritanceResetCommand_instances, _SpoFileRoleInheritanceResetCommand_initTelemetry, _SpoFileRoleInheritanceResetCommand_initOptions, _SpoFileRoleInheritanceResetCommand_initValidators, _SpoFileRoleInheritanceResetCommand_initOptionSets, _SpoFileRoleInheritanceResetCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFileRoleInheritanceResetCommand extends SpoCommand {
    get name() {
        return commands.FILE_ROLEINHERITANCE_RESET;
    }
    get description() {
        return 'Restores the role inheritance of a file';
    }
    constructor() {
        super();
        _SpoFileRoleInheritanceResetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFileRoleInheritanceResetCommand_instances, "m", _SpoFileRoleInheritanceResetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFileRoleInheritanceResetCommand_instances, "m", _SpoFileRoleInheritanceResetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFileRoleInheritanceResetCommand_instances, "m", _SpoFileRoleInheritanceResetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFileRoleInheritanceResetCommand_instances, "m", _SpoFileRoleInheritanceResetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFileRoleInheritanceResetCommand_instances, "m", _SpoFileRoleInheritanceResetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const resetFileRoleInheritance = async () => {
            if (this.verbose) {
                await logger.logToStderr(`Resetting role inheritance for file ${args.options.fileId || args.options.fileUrl}`);
            }
            try {
                const fileURL = await this.getFileURL(args, logger);
                const requestOptions = {
                    url: `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(fileURL)}')/ListItemAllFields/resetroleinheritance`,
                    headers: {
                        accept: 'application/json;odata=nometadata'
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
            await resetFileRoleInheritance();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to reset the role inheritance of file ${args.options.fileUrl || args.options.fileId} located in site ${args.options.webUrl}?` });
            if (result) {
                await resetFileRoleInheritance();
            }
        }
    }
    async getFileURL(args, logger) {
        if (args.options.fileUrl) {
            return urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl);
        }
        const file = await spo.getFileById(args.options.webUrl, args.options.fileId, logger, this.verbose);
        return file.ServerRelativeUrl;
    }
}
_SpoFileRoleInheritanceResetCommand_instances = new WeakSet(), _SpoFileRoleInheritanceResetCommand_initTelemetry = function _SpoFileRoleInheritanceResetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            fileUrl: typeof args.options.fileUrl !== 'undefined',
            fileId: typeof args.options.fileId !== 'undefined',
            force: !!args.options.force
        });
    });
}, _SpoFileRoleInheritanceResetCommand_initOptions = function _SpoFileRoleInheritanceResetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--fileUrl [fileUrl]'
    }, {
        option: '-i, --fileId [fileId]'
    }, {
        option: '-f, --force'
    });
}, _SpoFileRoleInheritanceResetCommand_initValidators = function _SpoFileRoleInheritanceResetCommand_initValidators() {
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
}, _SpoFileRoleInheritanceResetCommand_initOptionSets = function _SpoFileRoleInheritanceResetCommand_initOptionSets() {
    this.optionSets.push({ options: ['fileId', 'fileUrl'] });
}, _SpoFileRoleInheritanceResetCommand_initTypes = function _SpoFileRoleInheritanceResetCommand_initTypes() {
    this.types.string.push('webUrl', 'fileUrl', 'fileId');
    this.types.boolean.push('force');
};
export default new SpoFileRoleInheritanceResetCommand();
//# sourceMappingURL=file-roleinheritance-reset.js.map