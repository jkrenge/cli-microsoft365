var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFileSharingLinkGetCommand_instances, _SpoFileSharingLinkGetCommand_initTelemetry, _SpoFileSharingLinkGetCommand_initOptions, _SpoFileSharingLinkGetCommand_initValidators, _SpoFileSharingLinkGetCommand_initOptionSets, _SpoFileSharingLinkGetCommand_initTypes;
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFileSharingLinkGetCommand extends SpoCommand {
    get name() {
        return commands.FILE_SHARINGLINK_GET;
    }
    get description() {
        return 'Gets details about a specific sharing link of a file';
    }
    constructor() {
        super();
        _SpoFileSharingLinkGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFileSharingLinkGetCommand_instances, "m", _SpoFileSharingLinkGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFileSharingLinkGetCommand_instances, "m", _SpoFileSharingLinkGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFileSharingLinkGetCommand_instances, "m", _SpoFileSharingLinkGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFileSharingLinkGetCommand_instances, "m", _SpoFileSharingLinkGetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFileSharingLinkGetCommand_instances, "m", _SpoFileSharingLinkGetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving sharing link for file ${args.options.fileUrl || args.options.fileId} with id ${args.options.id}...`);
        }
        try {
            const fileDetails = await spo.getVroomFileDetails(args.options.webUrl, args.options.fileId, args.options.fileUrl);
            const requestOptions = {
                url: `https://graph.microsoft.com/v1.0/sites/${fileDetails.SiteId}/drives/${fileDetails.VroomDriveID}/items/${fileDetails.VroomItemID}/permissions/${args.options.id}`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            const res = await request.get(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoFileSharingLinkGetCommand_instances = new WeakSet(), _SpoFileSharingLinkGetCommand_initTelemetry = function _SpoFileSharingLinkGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            fileUrl: typeof args.options.fileUrl !== 'undefined',
            fileId: typeof args.options.fileId !== 'undefined'
        });
    });
}, _SpoFileSharingLinkGetCommand_initOptions = function _SpoFileSharingLinkGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--fileUrl [fileUrl]'
    }, {
        option: '--fileId [fileId]'
    }, {
        option: '-i, --id <id>'
    });
}, _SpoFileSharingLinkGetCommand_initValidators = function _SpoFileSharingLinkGetCommand_initValidators() {
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
}, _SpoFileSharingLinkGetCommand_initOptionSets = function _SpoFileSharingLinkGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['fileUrl', 'fileId'] });
}, _SpoFileSharingLinkGetCommand_initTypes = function _SpoFileSharingLinkGetCommand_initTypes() {
    this.types.string.push('webUrl', 'fileUrl', 'fileId', 'id');
};
export default new SpoFileSharingLinkGetCommand();
//# sourceMappingURL=file-sharinglink-get.js.map