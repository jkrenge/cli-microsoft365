var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFileRoleInheritanceBreakCommand_instances, _SpoFileRoleInheritanceBreakCommand_initTelemetry, _SpoFileRoleInheritanceBreakCommand_initOptions, _SpoFileRoleInheritanceBreakCommand_initValidators, _SpoFileRoleInheritanceBreakCommand_initOptionSets, _SpoFileRoleInheritanceBreakCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFileRoleInheritanceBreakCommand extends SpoCommand {
    get name() {
        return commands.FILE_ROLEINHERITANCE_BREAK;
    }
    get description() {
        return 'Breaks inheritance of a file. Keeping existing permissions is the default behavior.';
    }
    constructor() {
        super();
        _SpoFileRoleInheritanceBreakCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFileRoleInheritanceBreakCommand_instances, "m", _SpoFileRoleInheritanceBreakCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFileRoleInheritanceBreakCommand_instances, "m", _SpoFileRoleInheritanceBreakCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFileRoleInheritanceBreakCommand_instances, "m", _SpoFileRoleInheritanceBreakCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFileRoleInheritanceBreakCommand_instances, "m", _SpoFileRoleInheritanceBreakCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFileRoleInheritanceBreakCommand_instances, "m", _SpoFileRoleInheritanceBreakCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const breakFileRoleInheritance = async () => {
            if (this.verbose) {
                await logger.logToStderr(`Breaking role inheritance for file ${args.options.fileId || args.options.fileUrl}`);
            }
            try {
                const fileURL = await this.getFileURL(args, logger);
                const keepExistingPermissions = !args.options.clearExistingPermissions;
                const requestOptions = {
                    url: `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(fileURL)}')/ListItemAllFields/breakroleinheritance(${keepExistingPermissions})`,
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
            await breakFileRoleInheritance();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to break the role inheritance of file ${args.options.fileUrl || args.options.fileId} located in site ${args.options.webUrl}?` });
            if (result) {
                await breakFileRoleInheritance();
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
_SpoFileRoleInheritanceBreakCommand_instances = new WeakSet(), _SpoFileRoleInheritanceBreakCommand_initTelemetry = function _SpoFileRoleInheritanceBreakCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            fileUrl: typeof args.options.fileUrl !== 'undefined',
            fileId: typeof args.options.fileId !== 'undefined',
            clearExistingPermissions: !!args.options.clearExistingPermissions,
            force: !!args.options.force
        });
    });
}, _SpoFileRoleInheritanceBreakCommand_initOptions = function _SpoFileRoleInheritanceBreakCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--fileUrl [fileUrl]'
    }, {
        option: '-i, --fileId [fileId]'
    }, {
        option: '-c, --clearExistingPermissions'
    }, {
        option: '-f, --force'
    });
}, _SpoFileRoleInheritanceBreakCommand_initValidators = function _SpoFileRoleInheritanceBreakCommand_initValidators() {
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
}, _SpoFileRoleInheritanceBreakCommand_initOptionSets = function _SpoFileRoleInheritanceBreakCommand_initOptionSets() {
    this.optionSets.push({ options: ['fileId', 'fileUrl'] });
}, _SpoFileRoleInheritanceBreakCommand_initTypes = function _SpoFileRoleInheritanceBreakCommand_initTypes() {
    this.types.string.push('webUrl', 'fileUrl', 'fileId');
    this.types.boolean.push('clearExistingPermissions', 'force');
};
export default new SpoFileRoleInheritanceBreakCommand();
//# sourceMappingURL=file-roleinheritance-break.js.map