var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFileSharingLinkListCommand_instances, _a, _SpoFileSharingLinkListCommand_initTelemetry, _SpoFileSharingLinkListCommand_initOptions, _SpoFileSharingLinkListCommand_initValidators, _SpoFileSharingLinkListCommand_initOptionSets, _SpoFileSharingLinkListCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFileSharingLinkListCommand extends SpoCommand {
    get name() {
        return commands.FILE_SHARINGLINK_LIST;
    }
    get description() {
        return 'Lists all the sharing links of a specific file';
    }
    defaultProperties() {
        return ['id', 'scope', 'roles', 'link'];
    }
    constructor() {
        super();
        _SpoFileSharingLinkListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFileSharingLinkListCommand_instances, "m", _SpoFileSharingLinkListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFileSharingLinkListCommand_instances, "m", _SpoFileSharingLinkListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFileSharingLinkListCommand_instances, "m", _SpoFileSharingLinkListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFileSharingLinkListCommand_instances, "m", _SpoFileSharingLinkListCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFileSharingLinkListCommand_instances, "m", _SpoFileSharingLinkListCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving sharing links for file ${args.options.fileId || args.options.fileUrl}...`);
        }
        try {
            const fileDetails = await spo.getVroomFileDetails(args.options.webUrl, args.options.fileId, args.options.fileUrl);
            let url = `https://graph.microsoft.com/v1.0/sites/${fileDetails.SiteId}/drives/${fileDetails.VroomDriveID}/items/${fileDetails.VroomItemID}/permissions?$filter=Link ne null`;
            if (args.options.scope) {
                url += ` and Link/Scope eq '${args.options.scope}'`;
            }
            const sharingLinks = await odata.getAllItems(url);
            if (!args.options.output || !cli.shouldTrimOutput(args.options.output)) {
                await logger.log(sharingLinks);
            }
            else {
                //converted to text friendly output
                await logger.log(sharingLinks.map(i => {
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
_a = SpoFileSharingLinkListCommand, _SpoFileSharingLinkListCommand_instances = new WeakSet(), _SpoFileSharingLinkListCommand_initTelemetry = function _SpoFileSharingLinkListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            fileId: typeof args.options.fileId !== 'undefined',
            fileUrl: typeof args.options.fileUrl !== 'undefined',
            scope: typeof args.options.scope !== 'undefined'
        });
    });
}, _SpoFileSharingLinkListCommand_initOptions = function _SpoFileSharingLinkListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--fileId [fileId]'
    }, {
        option: '--fileUrl [fileUrl]'
    }, {
        option: '-s, --scope [scope]',
        autocomplete: _a.allowedScopes
    });
}, _SpoFileSharingLinkListCommand_initValidators = function _SpoFileSharingLinkListCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.fileId && !validation.isValidGuid(args.options.fileId)) {
            return `${args.options.fileId} is not a valid GUID`;
        }
        if (args.options.scope && _a.allowedScopes.indexOf(args.options.scope) === -1) {
            return `'${args.options.scope}' is not a valid scope. Allowed values are: ${_a.allowedScopes.join(',')}`;
        }
        return true;
    });
}, _SpoFileSharingLinkListCommand_initOptionSets = function _SpoFileSharingLinkListCommand_initOptionSets() {
    this.optionSets.push({ options: ['fileId', 'fileUrl'] });
}, _SpoFileSharingLinkListCommand_initTypes = function _SpoFileSharingLinkListCommand_initTypes() {
    this.types.string.push('webUrl', 'fileId', 'fileUrl', 'scope');
};
SpoFileSharingLinkListCommand.allowedScopes = ['anonymous', 'users', 'organization'];
export default new SpoFileSharingLinkListCommand();
//# sourceMappingURL=file-sharinglink-list.js.map