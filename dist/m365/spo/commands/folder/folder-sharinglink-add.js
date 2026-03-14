var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFolderSharingLinkAddCommand_instances, _SpoFolderSharingLinkAddCommand_initTelemetry, _SpoFolderSharingLinkAddCommand_initOptions, _SpoFolderSharingLinkAddCommand_initValidators, _SpoFolderSharingLinkAddCommand_initOptionSets, _SpoFolderSharingLinkAddCommand_initTypes;
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { drive } from '../../../../utils/drive.js';
import { validation } from '../../../../utils/validation.js';
import { formatting } from '../../../../utils/formatting.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFolderSharingLinkAddCommand extends SpoCommand {
    get name() {
        return commands.FOLDER_SHARINGLINK_ADD;
    }
    get description() {
        return 'Creates a new sharing link to a folder';
    }
    constructor() {
        super();
        _SpoFolderSharingLinkAddCommand_instances.add(this);
        this.allowedTypes = ['view', 'edit'];
        this.allowedScopes = ['anonymous', 'organization', 'users'];
        __classPrivateFieldGet(this, _SpoFolderSharingLinkAddCommand_instances, "m", _SpoFolderSharingLinkAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFolderSharingLinkAddCommand_instances, "m", _SpoFolderSharingLinkAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFolderSharingLinkAddCommand_instances, "m", _SpoFolderSharingLinkAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFolderSharingLinkAddCommand_instances, "m", _SpoFolderSharingLinkAddCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFolderSharingLinkAddCommand_instances, "m", _SpoFolderSharingLinkAddCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Creating a sharing link to a folder ${args.options.folderId || args.options.folderUrl}...`);
        }
        try {
            const relFolderUrl = await spo.getFolderServerRelativeUrl(args.options.webUrl, args.options.folderUrl, args.options.folderId, logger, args.options.verbose);
            const absoluteFolderUrl = urlUtil.getAbsoluteUrl(args.options.webUrl, relFolderUrl);
            const folderUrl = new URL(absoluteFolderUrl);
            const siteId = await spo.getSiteIdByMSGraph(args.options.webUrl);
            const driveDetails = await drive.getDriveByUrl(siteId, folderUrl, logger, args.options.verbose);
            const itemId = await drive.getDriveItemId(driveDetails, folderUrl, logger, args.options.verbose);
            const requestOptions = {
                url: `https://graph.microsoft.com/v1.0/drives/${driveDetails.id}/items/${itemId}/createLink`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: {
                    type: args.options.type,
                    expirationDateTime: args.options.expirationDateTime,
                    scope: args.options.scope,
                    retainInheritedPermissions: !!args.options.retainInheritiedPermissions
                }
            };
            if (args.options.scope === 'users' && args.options.recipients) {
                const recipients = formatting.splitAndTrim(args.options.recipients).map(email => ({ email }));
                requestOptions.data.recipients = recipients;
            }
            const sharingLink = await request.post(requestOptions);
            // remove grantedToIdentities from the sharing link object
            if (sharingLink.grantedToIdentities) {
                delete sharingLink.grantedToIdentities;
            }
            await logger.log(sharingLink);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoFolderSharingLinkAddCommand_instances = new WeakSet(), _SpoFolderSharingLinkAddCommand_initTelemetry = function _SpoFolderSharingLinkAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            folderId: typeof args.options.folderId !== 'undefined',
            folderUrl: typeof args.options.folderUrl !== 'undefined',
            type: typeof args.options.type !== 'undefined',
            expirationDateTime: typeof args.options.expirationDateTime !== 'undefined',
            scope: typeof args.options.scope !== 'undefined',
            retainInheritiedPermissions: !!args.options.retainInheritiedPermissions,
            recipients: typeof args.options.recipients !== 'undefined'
        });
    });
}, _SpoFolderSharingLinkAddCommand_initOptions = function _SpoFolderSharingLinkAddCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--folderId [folderId]'
    }, {
        option: '--folderUrl [folderUrl]'
    }, {
        option: '--type <type>',
        autocomplete: this.allowedTypes
    }, {
        option: '--expirationDateTime [expirationDateTime]'
    }, {
        option: '--scope [scope]',
        autocomplete: this.allowedScopes
    }, {
        option: '--retainInheritedPermissions [retainInheritedPermissions]'
    }, {
        option: '--recipients [recipients]'
    });
}, _SpoFolderSharingLinkAddCommand_initValidators = function _SpoFolderSharingLinkAddCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.folderId && !validation.isValidGuid(args.options.folderId)) {
            return `${args.options.folderId} is not a valid GUID`;
        }
        if (args.options.type && !this.allowedTypes.some(type => type === args.options.type)) {
            return `'${args.options.type}' is not a valid type. Allowed values are: ${this.allowedTypes.join(',')}`;
        }
        if (args.options.scope) {
            if (!this.allowedScopes.includes(args.options.scope)) {
                return `'${args.options.scope}' is not a valid scope. Allowed values are: ${this.allowedScopes.join(', ')}.`;
            }
        }
        if (args.options.expirationDateTime && !validation.isValidISODateTime(args.options.expirationDateTime)) {
            return `${args.options.expirationDateTime} is not a valid ISO date string.`;
        }
        if (args.options.recipients) {
            const isValidUPNArrayResult = validation.isValidUserPrincipalNameArray(args.options.recipients);
            if (isValidUPNArrayResult !== true) {
                return `The following user principal names are invalid for the option 'recipients': ${isValidUPNArrayResult}.`;
            }
        }
        return true;
    });
}, _SpoFolderSharingLinkAddCommand_initOptionSets = function _SpoFolderSharingLinkAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['folderId', 'folderUrl'] }, {
        options: ['recipients'],
        runsWhen: (args) => args.options.scope !== undefined && args.options.scope === 'users'
    });
}, _SpoFolderSharingLinkAddCommand_initTypes = function _SpoFolderSharingLinkAddCommand_initTypes() {
    this.types.string.push('webUrl', 'folderId', 'folderUrl', 'type', 'expirationDateTime', 'scope', 'recipients');
    this.types.boolean.push('retainInheritiedPermissions');
};
export default new SpoFolderSharingLinkAddCommand();
//# sourceMappingURL=folder-sharinglink-add.js.map