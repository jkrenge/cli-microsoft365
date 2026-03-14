var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFolderSharingLinkListCommand_instances, _SpoFolderSharingLinkListCommand_initTelemetry, _SpoFolderSharingLinkListCommand_initOptions, _SpoFolderSharingLinkListCommand_initOptionSets, _SpoFolderSharingLinkListCommand_initValidators, _SpoFolderSharingLinkListCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import { spo } from '../../../../utils/spo.js';
import { odata } from '../../../../utils/odata.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { drive } from '../../../../utils/drive.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFolderSharingLinkListCommand extends SpoCommand {
    get name() {
        return commands.FOLDER_SHARINGLINK_LIST;
    }
    get description() {
        return 'Lists all the sharing links of a specific folder';
    }
    defaultProperties() {
        return ['id', 'scope', 'roles', 'link'];
    }
    constructor() {
        super();
        _SpoFolderSharingLinkListCommand_instances.add(this);
        this.allowedScopes = ['anonymous', 'users', 'organization'];
        __classPrivateFieldGet(this, _SpoFolderSharingLinkListCommand_instances, "m", _SpoFolderSharingLinkListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFolderSharingLinkListCommand_instances, "m", _SpoFolderSharingLinkListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFolderSharingLinkListCommand_instances, "m", _SpoFolderSharingLinkListCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFolderSharingLinkListCommand_instances, "m", _SpoFolderSharingLinkListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFolderSharingLinkListCommand_instances, "m", _SpoFolderSharingLinkListCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving sharing links for folder ${args.options.folderId || args.options.folderUrl}...`);
        }
        try {
            const relFolderUrl = await spo.getFolderServerRelativeUrl(args.options.webUrl, args.options.folderUrl, args.options.folderId, logger, args.options.verbose);
            const absoluteFolderUrl = urlUtil.getAbsoluteUrl(args.options.webUrl, relFolderUrl);
            const folderUrl = new URL(absoluteFolderUrl);
            const siteId = await spo.getSiteIdByMSGraph(args.options.webUrl);
            const driveDetails = await drive.getDriveByUrl(siteId, folderUrl, logger, args.options.verbose);
            const itemId = await drive.getDriveItemId(driveDetails, folderUrl, logger, args.options.verbose);
            let requestUrl = `https://graph.microsoft.com/v1.0/drives/${driveDetails.id}/items/${itemId}/permissions?$filter=Link ne null`;
            if (args.options.scope) {
                requestUrl += ` and Link/Scope eq '${args.options.scope}'`;
            }
            const sharingLinks = await odata.getAllItems(requestUrl);
            // remove grantedToIdentities from the sharing link object
            const filteredSharingLinks = sharingLinks.map(link => {
                // eslint-disable-next-line @typescript-eslint/no-unused-vars
                const { grantedToIdentities, ...filteredLink } = link;
                return filteredLink;
            });
            if (!args.options.output || !cli.shouldTrimOutput(args.options.output)) {
                await logger.log(filteredSharingLinks);
            }
            else {
                //converted to text friendly output
                await logger.log(filteredSharingLinks.map(i => {
                    return {
                        id: i.id,
                        roles: i.roles.join(','),
                        link: i.link.webUrl,
                        scope: i.link.scope
                    };
                }));
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoFolderSharingLinkListCommand_instances = new WeakSet(), _SpoFolderSharingLinkListCommand_initTelemetry = function _SpoFolderSharingLinkListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            folderUrl: typeof args.options.folderUrl !== 'undefined',
            folderId: typeof args.options.folderId !== 'undefined',
            scope: typeof args.options.scope !== 'undefined'
        });
    });
}, _SpoFolderSharingLinkListCommand_initOptions = function _SpoFolderSharingLinkListCommand_initOptions() {
    this.options.unshift({ option: '-u, --webUrl <webUrl>' }, { option: '--folderUrl [folderUrl]' }, { option: '--folderId [folderId]' }, {
        option: '-s, --scope [scope]',
        autocomplete: this.allowedScopes
    });
}, _SpoFolderSharingLinkListCommand_initOptionSets = function _SpoFolderSharingLinkListCommand_initOptionSets() {
    this.optionSets.push({ options: ['folderUrl', 'folderId'] });
}, _SpoFolderSharingLinkListCommand_initValidators = function _SpoFolderSharingLinkListCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.folderId && !validation.isValidGuid(args.options.folderId)) {
            return `${args.options.folderId} is not a valid GUID`;
        }
        if (args.options.scope && !this.allowedScopes.some(scope => scope === args.options.scope)) {
            return `'${args.options.scope}' is not a valid scope. Allowed values are: ${this.allowedScopes.join(',')}`;
        }
        return true;
    });
}, _SpoFolderSharingLinkListCommand_initTypes = function _SpoFolderSharingLinkListCommand_initTypes() {
    this.types.string.push('webUrl', 'folderUrl', 'folderId', 'scope');
};
export default new SpoFolderSharingLinkListCommand();
//# sourceMappingURL=folder-sharinglink-list.js.map