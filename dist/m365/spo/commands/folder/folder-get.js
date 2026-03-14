var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFolderGetCommand_instances, _SpoFolderGetCommand_initTelemetry, _SpoFolderGetCommand_initOptions, _SpoFolderGetCommand_initValidators, _SpoFolderGetCommand_initOptionSets, _SpoFolderGetCommand_initTypes;
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { ListPrincipalType } from '../list/ListPrincipalType.js';
class SpoFolderGetCommand extends SpoCommand {
    get name() {
        return commands.FOLDER_GET;
    }
    get description() {
        return 'Gets information about the specified folder';
    }
    constructor() {
        super();
        _SpoFolderGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFolderGetCommand_instances, "m", _SpoFolderGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFolderGetCommand_instances, "m", _SpoFolderGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFolderGetCommand_instances, "m", _SpoFolderGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFolderGetCommand_instances, "m", _SpoFolderGetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFolderGetCommand_instances, "m", _SpoFolderGetCommand_initTypes).call(this);
    }
    getExcludedOptionsWithUrls() {
        return ['url'];
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving folder from site ${args.options.webUrl}...`);
        }
        let requestUrl = `${args.options.webUrl}/_api/web`;
        if (args.options.id) {
            requestUrl += `/GetFolderById('${formatting.encodeQueryParameter(args.options.id)}')`;
        }
        else if (args.options.url) {
            const serverRelativePath = urlUtil.getServerRelativePath(args.options.webUrl, args.options.url);
            requestUrl += `/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')`;
        }
        if (args.options.withPermissions) {
            requestUrl += `?$expand=ListItemAllFields/HasUniqueRoleAssignments,ListItemAllFields/RoleAssignments/Member,ListItemAllFields/RoleAssignments/RoleDefinitionBindings`;
        }
        const requestOptions = {
            url: requestUrl,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        try {
            const folder = await request.get(requestOptions);
            if (args.options.withPermissions) {
                const listItemAllFields = folder.ListItemAllFields;
                if (!(listItemAllFields ?? false)) {
                    throw Error('Please ensure the specified folder URL or folder Id does not refer to a root folder. Use \'spo list get\' with withPermissions instead.');
                }
                listItemAllFields.RoleAssignments.forEach(r => {
                    r.Member.PrincipalTypeString = ListPrincipalType[r.Member.PrincipalType];
                    r.RoleDefinitionBindings = formatting.setFriendlyPermissions(r.RoleDefinitionBindings);
                });
            }
            await logger.log(folder);
        }
        catch (err) {
            if (err.statusCode && err.statusCode === 500) {
                throw new CommandError('Please check the folder URL. Folder might not exist on the specified URL');
            }
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoFolderGetCommand_instances = new WeakSet(), _SpoFolderGetCommand_initTelemetry = function _SpoFolderGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            url: typeof args.options.url !== 'undefined',
            withPermissions: typeof args.options.withPermissions !== 'undefined'
        });
    });
}, _SpoFolderGetCommand_initOptions = function _SpoFolderGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--url [url]'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '--withPermissions'
    });
}, _SpoFolderGetCommand_initValidators = function _SpoFolderGetCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.id) {
            if (!validation.isValidGuid(args.options.id)) {
                return `${args.options.id} is not a valid GUID`;
            }
        }
        return true;
    });
}, _SpoFolderGetCommand_initOptionSets = function _SpoFolderGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['url', 'id'] });
}, _SpoFolderGetCommand_initTypes = function _SpoFolderGetCommand_initTypes() {
    this.types.string.push('webUrl', 'url', 'id');
};
export default new SpoFolderGetCommand();
//# sourceMappingURL=folder-get.js.map