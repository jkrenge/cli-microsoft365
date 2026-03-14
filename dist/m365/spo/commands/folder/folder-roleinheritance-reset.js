var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFolderRoleInheritanceResetCommand_instances, _SpoFolderRoleInheritanceResetCommand_initTelemetry, _SpoFolderRoleInheritanceResetCommand_initOptions, _SpoFolderRoleInheritanceResetCommand_initValidators, _SpoFolderRoleInheritanceResetCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFolderRoleInheritanceResetCommand extends SpoCommand {
    get name() {
        return commands.FOLDER_ROLEINHERITANCE_RESET;
    }
    get description() {
        return 'Restores the role inheritance of a folder';
    }
    constructor() {
        super();
        _SpoFolderRoleInheritanceResetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFolderRoleInheritanceResetCommand_instances, "m", _SpoFolderRoleInheritanceResetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFolderRoleInheritanceResetCommand_instances, "m", _SpoFolderRoleInheritanceResetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFolderRoleInheritanceResetCommand_instances, "m", _SpoFolderRoleInheritanceResetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFolderRoleInheritanceResetCommand_instances, "m", _SpoFolderRoleInheritanceResetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const serverRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.folderUrl);
        const roleFolderUrl = urlUtil.getWebRelativePath(args.options.webUrl, args.options.folderUrl);
        let requestUrl = `${args.options.webUrl}/_api/web/`;
        const resetFolderRoleInheritance = async () => {
            try {
                if (roleFolderUrl.split('/').length === 2) {
                    requestUrl += `GetList('${formatting.encodeQueryParameter(serverRelativeUrl)}')`;
                }
                else {
                    requestUrl += `GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}')/ListItemAllFields`;
                }
                const requestOptions = {
                    url: `${requestUrl}/resetroleinheritance`,
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
            await resetFolderRoleInheritance();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to reset the role inheritance of folder ${args.options.folderUrl} located in site ${args.options.webUrl}?` });
            if (result) {
                await resetFolderRoleInheritance();
            }
        }
    }
}
_SpoFolderRoleInheritanceResetCommand_instances = new WeakSet(), _SpoFolderRoleInheritanceResetCommand_initTelemetry = function _SpoFolderRoleInheritanceResetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: !!args.options.force
        });
    });
}, _SpoFolderRoleInheritanceResetCommand_initOptions = function _SpoFolderRoleInheritanceResetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--folderUrl <folderUrl>'
    }, {
        option: '-f, --force'
    });
}, _SpoFolderRoleInheritanceResetCommand_initValidators = function _SpoFolderRoleInheritanceResetCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.webUrl));
}, _SpoFolderRoleInheritanceResetCommand_initTypes = function _SpoFolderRoleInheritanceResetCommand_initTypes() {
    this.types.string.push('webUrl', 'folderUrl');
    this.types.boolean.push('force');
};
export default new SpoFolderRoleInheritanceResetCommand();
//# sourceMappingURL=folder-roleinheritance-reset.js.map