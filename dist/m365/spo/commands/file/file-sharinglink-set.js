var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFileSharingLinkSetCommand_instances, _SpoFileSharingLinkSetCommand_initTelemetry, _SpoFileSharingLinkSetCommand_initOptions, _SpoFileSharingLinkSetCommand_initValidators, _SpoFileSharingLinkSetCommand_initOptionSets, _SpoFileSharingLinkSetCommand_initTypes;
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFileSharingLinkSetCommand extends SpoCommand {
    get name() {
        return commands.FILE_SHARINGLINK_SET;
    }
    get description() {
        return 'Updates a sharing link of a file';
    }
    constructor() {
        super();
        _SpoFileSharingLinkSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFileSharingLinkSetCommand_instances, "m", _SpoFileSharingLinkSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFileSharingLinkSetCommand_instances, "m", _SpoFileSharingLinkSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFileSharingLinkSetCommand_instances, "m", _SpoFileSharingLinkSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFileSharingLinkSetCommand_instances, "m", _SpoFileSharingLinkSetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFileSharingLinkSetCommand_instances, "m", _SpoFileSharingLinkSetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Updating sharing link of file ${args.options.fileId || args.options.fileUrl}...`);
        }
        try {
            const fileDetails = await spo.getVroomFileDetails(args.options.webUrl, args.options.fileId, args.options.fileUrl);
            const requestOptions = {
                url: `https://graph.microsoft.com/v1.0/sites/${fileDetails.SiteId}/drives/${fileDetails.VroomDriveID}/items/${fileDetails.VroomItemID}/permissions/${args.options.id}`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: {
                    expirationDateTime: args.options.expirationDateTime
                }
            };
            const sharingLink = await request.patch(requestOptions);
            await logger.log(sharingLink);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoFileSharingLinkSetCommand_instances = new WeakSet(), _SpoFileSharingLinkSetCommand_initTelemetry = function _SpoFileSharingLinkSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            fileId: typeof args.options.fileId !== 'undefined',
            fileUrl: typeof args.options.fileUrl !== 'undefined'
        });
    });
}, _SpoFileSharingLinkSetCommand_initOptions = function _SpoFileSharingLinkSetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--fileId [fileId]'
    }, {
        option: '--fileUrl [fileUrl]'
    }, {
        option: '--id <id>'
    }, {
        option: '--expirationDateTime <expirationDateTime>'
    });
}, _SpoFileSharingLinkSetCommand_initValidators = function _SpoFileSharingLinkSetCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.fileId && !validation.isValidGuid(args.options.fileId)) {
            return `${args.options.fileId} is not a valid GUID`;
        }
        if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        if (!validation.isValidISODateTime(args.options.expirationDateTime)) {
            return `'${args.options.expirationDateTime}' is not a valid ISO date string`;
        }
        return true;
    });
}, _SpoFileSharingLinkSetCommand_initOptionSets = function _SpoFileSharingLinkSetCommand_initOptionSets() {
    this.optionSets.push({ options: ['fileId', 'fileUrl'] });
}, _SpoFileSharingLinkSetCommand_initTypes = function _SpoFileSharingLinkSetCommand_initTypes() {
    this.types.string.push('webUrl', 'id', 'fileId', 'fileUrl', 'expirationDateTime');
};
export default new SpoFileSharingLinkSetCommand();
//# sourceMappingURL=file-sharinglink-set.js.map