var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFolderRoleInheritanceBreakCommand_instances, _SpoFolderRoleInheritanceBreakCommand_initTelemetry, _SpoFolderRoleInheritanceBreakCommand_initOptions, _SpoFolderRoleInheritanceBreakCommand_initValidators, _SpoFolderRoleInheritanceBreakCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFolderRoleInheritanceBreakCommand extends SpoCommand {
    get name() {
        return commands.FOLDER_ROLEINHERITANCE_BREAK;
    }
    get description() {
        return 'Breaks the role inheritance of a folder. Keeping existing permissions is the default behavior.';
    }
    constructor() {
        super();
        _SpoFolderRoleInheritanceBreakCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFolderRoleInheritanceBreakCommand_instances, "m", _SpoFolderRoleInheritanceBreakCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFolderRoleInheritanceBreakCommand_instances, "m", _SpoFolderRoleInheritanceBreakCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFolderRoleInheritanceBreakCommand_instances, "m", _SpoFolderRoleInheritanceBreakCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFolderRoleInheritanceBreakCommand_instances, "m", _SpoFolderRoleInheritanceBreakCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const keepExistingPermissions = !args.options.clearExistingPermissions;
        const serverRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.folderUrl);
        const roleFolderUrl = urlUtil.getWebRelativePath(args.options.webUrl, args.options.folderUrl);
        let requestUrl = `${args.options.webUrl}/_api/web/`;
        const breakFolderRoleInheritance = async () => {
            try {
                if (roleFolderUrl.split('/').length === 2) {
                    requestUrl += `GetList('${formatting.encodeQueryParameter(serverRelativeUrl)}')`;
                }
                else {
                    requestUrl += `GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}')/ListItemAllFields`;
                }
                const requestOptions = {
                    url: `${requestUrl}/breakroleinheritance(${keepExistingPermissions})`,
                    headers: {
                        accept: 'application/json'
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
            await breakFolderRoleInheritance();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to break the role inheritance of folder ${args.options.folderUrl} located in site ${args.options.webUrl}?` });
            if (result) {
                await breakFolderRoleInheritance();
            }
        }
    }
}
_SpoFolderRoleInheritanceBreakCommand_instances = new WeakSet(), _SpoFolderRoleInheritanceBreakCommand_initTelemetry = function _SpoFolderRoleInheritanceBreakCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            clearExistingPermissions: !!args.options.clearExistingPermissions,
            force: !!args.options.force
        });
    });
}, _SpoFolderRoleInheritanceBreakCommand_initOptions = function _SpoFolderRoleInheritanceBreakCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--folderUrl <folderUrl>'
    }, {
        option: '-c, --clearExistingPermissions'
    }, {
        option: '-f, --force'
    });
}, _SpoFolderRoleInheritanceBreakCommand_initValidators = function _SpoFolderRoleInheritanceBreakCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.webUrl));
}, _SpoFolderRoleInheritanceBreakCommand_initTypes = function _SpoFolderRoleInheritanceBreakCommand_initTypes() {
    this.types.string.push('webUrl', 'folderUrl');
    this.types.boolean.push('clearExistingPermissions', 'force');
};
export default new SpoFolderRoleInheritanceBreakCommand();
//# sourceMappingURL=folder-roleinheritance-break.js.map