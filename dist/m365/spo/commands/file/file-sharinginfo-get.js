var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFileSharingInfoGetCommand_instances, _SpoFileSharingInfoGetCommand_initTelemetry, _SpoFileSharingInfoGetCommand_initOptions, _SpoFileSharingInfoGetCommand_initValidators, _SpoFileSharingInfoGetCommand_initOptionSets, _SpoFileSharingInfoGetCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { FileSharingPrincipalType } from './FileSharingPrincipalType.js';
class SpoFileSharingInfoGetCommand extends SpoCommand {
    get name() {
        return commands.FILE_SHARINGINFO_GET;
    }
    get description() {
        return 'Generates a sharing information report for the specified file';
    }
    constructor() {
        super();
        _SpoFileSharingInfoGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFileSharingInfoGetCommand_instances, "m", _SpoFileSharingInfoGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFileSharingInfoGetCommand_instances, "m", _SpoFileSharingInfoGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFileSharingInfoGetCommand_instances, "m", _SpoFileSharingInfoGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFileSharingInfoGetCommand_instances, "m", _SpoFileSharingInfoGetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFileSharingInfoGetCommand_instances, "m", _SpoFileSharingInfoGetCommand_initTypes).call(this);
    }
    getExcludedOptionsWithUrls() {
        return ['fileUrl'];
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving sharing information report for the file ${args.options.fileId || args.options.fileUrl}`);
        }
        try {
            const fileInformation = await this.getNeededFileInformation(args);
            if (this.verbose) {
                await logger.logToStderr(`Retrieving sharing information report for the file with item Id  ${fileInformation.fileItemId}`);
            }
            const requestOptions = {
                url: `${args.options.webUrl}/_api/web/lists/getbytitle('${formatting.encodeQueryParameter(fileInformation.libraryName)}')/items(${fileInformation.fileItemId})/GetSharingInformation?$select=permissionsInformation&$Expand=permissionsInformation`,
                headers: {
                    'accept': 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            const res = await request.post(requestOptions);
            // typically, we don't do this, but in this case, we need to due to
            // the complexity of the retrieved object and the fact that we can't
            // use the generic way of simplifying the output
            if (!cli.shouldTrimOutput(args.options.output)) {
                await logger.log(res);
            }
            else {
                const fileSharingInfoCollection = [];
                res.permissionsInformation.links.forEach(link => {
                    link.linkDetails.Invitations.forEach(linkInvite => {
                        fileSharingInfoCollection.push({
                            SharedWith: linkInvite.invitee.name,
                            IsActive: linkInvite.invitee.isActive,
                            IsExternal: linkInvite.invitee.isExternal,
                            PrincipalType: FileSharingPrincipalType[parseInt(linkInvite.invitee.principalType)]
                        });
                    });
                });
                res.permissionsInformation.principals.forEach(principal => {
                    fileSharingInfoCollection.push({
                        SharedWith: principal.principal.name,
                        IsActive: principal.principal.isActive,
                        IsExternal: principal.principal.isExternal,
                        PrincipalType: FileSharingPrincipalType[parseInt(principal.principal.principalType)]
                    });
                });
                await logger.log(fileSharingInfoCollection);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getNeededFileInformation(args) {
        let requestUrl;
        if (args.options.fileId) {
            requestUrl = `${args.options.webUrl}/_api/web/GetFileById('${args.options.fileId}')/?$select=ListItemAllFields/Id,ListItemAllFields/ParentList/Title&$expand=ListItemAllFields/ParentList`;
        }
        else {
            const serverRelPath = urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl);
            requestUrl = `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(serverRelPath)}')?$select=ListItemAllFields/Id,ListItemAllFields/ParentList/Title&$expand=ListItemAllFields/ParentList`;
        }
        const requestOptions = {
            url: requestUrl,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const res = await request.get(requestOptions);
        return {
            fileItemId: parseInt(res.ListItemAllFields.Id),
            libraryName: res.ListItemAllFields.ParentList.Title
        };
    }
}
_SpoFileSharingInfoGetCommand_instances = new WeakSet(), _SpoFileSharingInfoGetCommand_initTelemetry = function _SpoFileSharingInfoGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            fileId: typeof args.options.fileId !== 'undefined',
            fileUrl: typeof args.options.fileUrl !== 'undefined'
        });
    });
}, _SpoFileSharingInfoGetCommand_initOptions = function _SpoFileSharingInfoGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-i, --fileId [fileId]'
    }, {
        option: '--fileUrl [fileUrl]'
    });
}, _SpoFileSharingInfoGetCommand_initValidators = function _SpoFileSharingInfoGetCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.fileId) {
            if (!validation.isValidGuid(args.options.fileId)) {
                return `${args.options.fileId} is not a valid GUID`;
            }
        }
        return true;
    });
}, _SpoFileSharingInfoGetCommand_initOptionSets = function _SpoFileSharingInfoGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['fileId', 'fileUrl'] });
}, _SpoFileSharingInfoGetCommand_initTypes = function _SpoFileSharingInfoGetCommand_initTypes() {
    this.types.string.push('webUrl', 'fileId', 'fileUrl');
};
export default new SpoFileSharingInfoGetCommand();
//# sourceMappingURL=file-sharinginfo-get.js.map