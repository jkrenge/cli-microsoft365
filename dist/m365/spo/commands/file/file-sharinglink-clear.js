var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFileSharingLinkClearCommand_instances, _SpoFileSharingLinkClearCommand_initTelemetry, _SpoFileSharingLinkClearCommand_initOptions, _SpoFileSharingLinkClearCommand_initValidators, _SpoFileSharingLinkClearCommand_initOptionSets, _SpoFileSharingLinkClearCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import spoFileSharingLinkListCommand from './file-sharinglink-list.js';
class SpoFileSharingLinkClearCommand extends SpoCommand {
    get name() {
        return commands.FILE_SHARINGLINK_CLEAR;
    }
    get description() {
        return 'Removes sharing links of a file';
    }
    constructor() {
        super();
        _SpoFileSharingLinkClearCommand_instances.add(this);
        this.allowedScopes = ['anonymous', 'users', 'organization'];
        __classPrivateFieldGet(this, _SpoFileSharingLinkClearCommand_instances, "m", _SpoFileSharingLinkClearCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFileSharingLinkClearCommand_instances, "m", _SpoFileSharingLinkClearCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFileSharingLinkClearCommand_instances, "m", _SpoFileSharingLinkClearCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFileSharingLinkClearCommand_instances, "m", _SpoFileSharingLinkClearCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFileSharingLinkClearCommand_instances, "m", _SpoFileSharingLinkClearCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const clearSharingLinks = async () => {
            try {
                if (this.verbose) {
                    await logger.logToStderr(`Clearing sharing links for file ${args.options.fileUrl || args.options.fileId}${args.options.scope ? ` with scope ${args.options.scope}` : ''}`);
                }
                const fileDetails = await spo.getVroomFileDetails(args.options.webUrl, args.options.fileId, args.options.fileUrl);
                const sharingLinks = await this.getFileSharingLinks(args.options.webUrl, args.options.fileId, args.options.fileUrl, args.options.scope);
                const requestOptions = {
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    },
                    responseType: 'json'
                };
                for (const sharingLink of sharingLinks) {
                    requestOptions.url = `https://graph.microsoft.com/v1.0/sites/${fileDetails.SiteId}/drives/${fileDetails.VroomDriveID}/items/${fileDetails.VroomItemID}/permissions/${sharingLink.id}`;
                    await request.delete(requestOptions);
                }
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await clearSharingLinks();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to clear the sharing links of file ${args.options.fileUrl || args.options.fileId}${args.options.scope ? ` with scope ${args.options.scope}` : ''}?` });
            if (result) {
                await clearSharingLinks();
            }
        }
    }
    async getFileSharingLinks(webUrl, fileId, fileUrl, scope) {
        const sharingLinkListOptions = {
            webUrl: webUrl,
            fileId: fileId,
            fileUrl: fileUrl,
            scope: scope,
            debug: this.debug,
            verbose: this.verbose
        };
        const commandOutput = await cli.executeCommandWithOutput(spoFileSharingLinkListCommand, { options: { ...sharingLinkListOptions, _: [] } });
        const outputParsed = JSON.parse(commandOutput.stdout);
        return outputParsed;
    }
}
_SpoFileSharingLinkClearCommand_instances = new WeakSet(), _SpoFileSharingLinkClearCommand_initTelemetry = function _SpoFileSharingLinkClearCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            fileUrl: typeof args.options.fileUrl !== 'undefined',
            fileId: typeof args.options.fileId !== 'undefined',
            scope: typeof args.options.scope !== 'undefined',
            force: !!args.options.force
        });
    });
}, _SpoFileSharingLinkClearCommand_initOptions = function _SpoFileSharingLinkClearCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--fileUrl [fileUrl]'
    }, {
        option: '--fileId [fileId]'
    }, {
        option: '-s, --scope [scope]',
        autocomplete: this.allowedScopes
    }, {
        option: '-f, --force'
    });
}, _SpoFileSharingLinkClearCommand_initValidators = function _SpoFileSharingLinkClearCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.fileId && !validation.isValidGuid(args.options.fileId)) {
            return `${args.options.fileId} is not a valid GUID`;
        }
        if (args.options.scope && this.allowedScopes.indexOf(args.options.scope) === -1) {
            return `'${args.options.scope}' is not a valid scope. Allowed values are: ${this.allowedScopes.join(',')}`;
        }
        return true;
    });
}, _SpoFileSharingLinkClearCommand_initOptionSets = function _SpoFileSharingLinkClearCommand_initOptionSets() {
    this.optionSets.push({ options: ['fileUrl', 'fileId'] });
}, _SpoFileSharingLinkClearCommand_initTypes = function _SpoFileSharingLinkClearCommand_initTypes() {
    this.types.string.push('webUrl', 'fileUrl', 'fileId', 'scope');
    this.types.boolean.push('force');
};
export default new SpoFileSharingLinkClearCommand();
//# sourceMappingURL=file-sharinglink-clear.js.map