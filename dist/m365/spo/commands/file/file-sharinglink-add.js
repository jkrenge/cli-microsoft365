var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFileSharingLinkAddCommand_instances, _a, _SpoFileSharingLinkAddCommand_initTelemetry, _SpoFileSharingLinkAddCommand_initOptions, _SpoFileSharingLinkAddCommand_initValidators, _SpoFileSharingLinkAddCommand_initOptionSets, _SpoFileSharingLinkAddCommand_initTypes;
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFileSharingLinkAddCommand extends SpoCommand {
    get name() {
        return commands.FILE_SHARINGLINK_ADD;
    }
    get description() {
        return 'Creates a new sharing link for a file';
    }
    constructor() {
        super();
        _SpoFileSharingLinkAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFileSharingLinkAddCommand_instances, "m", _SpoFileSharingLinkAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFileSharingLinkAddCommand_instances, "m", _SpoFileSharingLinkAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFileSharingLinkAddCommand_instances, "m", _SpoFileSharingLinkAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFileSharingLinkAddCommand_instances, "m", _SpoFileSharingLinkAddCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFileSharingLinkAddCommand_instances, "m", _SpoFileSharingLinkAddCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Creating a sharing link for file ${args.options.fileId || args.options.fileUrl}...`);
        }
        try {
            const fileDetails = await spo.getVroomFileDetails(args.options.webUrl, args.options.fileId, args.options.fileUrl);
            const requestOptions = {
                url: `https://graph.microsoft.com/v1.0/sites/${fileDetails.SiteId}/drives/${fileDetails.VroomDriveID}/items/${fileDetails.VroomItemID}/createLink`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: {
                    type: args.options.type,
                    expirationDateTime: args.options.expirationDateTime,
                    scope: args.options.scope
                }
            };
            const sharingLink = await request.post(requestOptions);
            await logger.log(sharingLink);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_a = SpoFileSharingLinkAddCommand, _SpoFileSharingLinkAddCommand_instances = new WeakSet(), _SpoFileSharingLinkAddCommand_initTelemetry = function _SpoFileSharingLinkAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            fileId: typeof args.options.fileId !== 'undefined',
            fileUrl: typeof args.options.fileUrl !== 'undefined',
            expirationDateTime: typeof args.options.expirationDateTime !== 'undefined',
            scope: typeof args.options.scope !== 'undefined'
        });
    });
}, _SpoFileSharingLinkAddCommand_initOptions = function _SpoFileSharingLinkAddCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--fileId [fileId]'
    }, {
        option: '--fileUrl [fileUrl]'
    }, {
        option: '--type <type>',
        autocomplete: _a.types
    }, {
        option: '--expirationDateTime [expirationDateTime]'
    }, {
        option: '--scope [scope]',
        autocomplete: _a.scopes
    });
}, _SpoFileSharingLinkAddCommand_initValidators = function _SpoFileSharingLinkAddCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.fileId && !validation.isValidGuid(args.options.fileId)) {
            return `${args.options.fileId} is not a valid GUID`;
        }
        if (_a.types.indexOf(args.options.type) < 0) {
            return `'${args.options.type}' is not a valid type. Allowed types are ${_a.types.join(', ')}`;
        }
        if (args.options.scope &&
            _a.scopes.indexOf(args.options.scope) < 0) {
            return `'${args.options.scope}' is not a valid scope. Allowed scopes are ${_a.scopes.join(', ')}`;
        }
        if (args.options.scope && args.options.scope !== 'anonymous' && args.options.expirationDateTime) {
            return `Option expirationDateTime can only be used for links with scope 'anonymous'.`;
        }
        if (args.options.expirationDateTime && !validation.isValidISODateTime(args.options.expirationDateTime)) {
            return `${args.options.expirationDateTime} is not a valid ISO date string.`;
        }
        return true;
    });
}, _SpoFileSharingLinkAddCommand_initOptionSets = function _SpoFileSharingLinkAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['fileId', 'fileUrl'] });
}, _SpoFileSharingLinkAddCommand_initTypes = function _SpoFileSharingLinkAddCommand_initTypes() {
    this.types.string.push('webUrl', 'fileId', 'fileUrl', 'type', 'expirationDateTime', 'scope');
};
SpoFileSharingLinkAddCommand.types = ['view', 'edit'];
SpoFileSharingLinkAddCommand.scopes = ['anonymous', 'organization'];
export default new SpoFileSharingLinkAddCommand();
//# sourceMappingURL=file-sharinglink-add.js.map