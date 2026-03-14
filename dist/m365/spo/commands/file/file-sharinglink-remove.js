var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFileSharingLinkRemoveCommand_instances, _SpoFileSharingLinkRemoveCommand_initTelemetry, _SpoFileSharingLinkRemoveCommand_initOptions, _SpoFileSharingLinkRemoveCommand_initValidators, _SpoFileSharingLinkRemoveCommand_initOptionSets, _SpoFileSharingLinkRemoveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFileSharingLinkRemoveCommand extends SpoCommand {
    get name() {
        return commands.FILE_SHARINGLINK_REMOVE;
    }
    get description() {
        return 'Removes a specific sharing link of a file';
    }
    constructor() {
        super();
        _SpoFileSharingLinkRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFileSharingLinkRemoveCommand_instances, "m", _SpoFileSharingLinkRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFileSharingLinkRemoveCommand_instances, "m", _SpoFileSharingLinkRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFileSharingLinkRemoveCommand_instances, "m", _SpoFileSharingLinkRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFileSharingLinkRemoveCommand_instances, "m", _SpoFileSharingLinkRemoveCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFileSharingLinkRemoveCommand_instances, "m", _SpoFileSharingLinkRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const removeSharingLink = async () => {
            try {
                if (this.verbose) {
                    await logger.logToStderr(`Removing sharing link of file ${args.options.fileUrl || args.options.fileId} with id ${args.options.id}...`);
                }
                const fileDetails = await spo.getVroomFileDetails(args.options.webUrl, args.options.fileId, args.options.fileUrl);
                const requestOptions = {
                    url: `https://graph.microsoft.com/v1.0/sites/${fileDetails.SiteId}/drives/${fileDetails.VroomDriveID}/items/${fileDetails.VroomItemID}/permissions/${args.options.id}`,
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    },
                    responseType: 'json'
                };
                await request.delete(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeSharingLink();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove sharing link ${args.options.id} of file ${args.options.fileUrl || args.options.fileId}?` });
            if (result) {
                await removeSharingLink();
            }
        }
    }
}
_SpoFileSharingLinkRemoveCommand_instances = new WeakSet(), _SpoFileSharingLinkRemoveCommand_initTelemetry = function _SpoFileSharingLinkRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            fileUrl: typeof args.options.fileUrl !== 'undefined',
            fileId: typeof args.options.fileId !== 'undefined',
            force: !!args.options.force
        });
    });
}, _SpoFileSharingLinkRemoveCommand_initOptions = function _SpoFileSharingLinkRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--fileUrl [fileUrl]'
    }, {
        option: '--fileId [fileId]'
    }, {
        option: '-i, --id <id>'
    }, {
        option: '-f, --force'
    });
}, _SpoFileSharingLinkRemoveCommand_initValidators = function _SpoFileSharingLinkRemoveCommand_initValidators() {
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
}, _SpoFileSharingLinkRemoveCommand_initOptionSets = function _SpoFileSharingLinkRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['fileUrl', 'fileId'] });
}, _SpoFileSharingLinkRemoveCommand_initTypes = function _SpoFileSharingLinkRemoveCommand_initTypes() {
    this.types.string.push('webUrl', 'fileUrl', 'fileId', 'id');
    this.types.boolean.push('force');
};
export default new SpoFileSharingLinkRemoveCommand();
//# sourceMappingURL=file-sharinglink-remove.js.map