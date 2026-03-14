var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFolderRoleAssignmentRemoveCommand_instances, _SpoFolderRoleAssignmentRemoveCommand_initTelemetry, _SpoFolderRoleAssignmentRemoveCommand_initOptions, _SpoFolderRoleAssignmentRemoveCommand_initValidators, _SpoFolderRoleAssignmentRemoveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { spo } from '../../../../utils/spo.js';
class SpoFolderRoleAssignmentRemoveCommand extends SpoCommand {
    get name() {
        return commands.FOLDER_ROLEASSIGNMENT_REMOVE;
    }
    get description() {
        return 'Removes a role assignment from the specified folder';
    }
    constructor() {
        super();
        _SpoFolderRoleAssignmentRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFolderRoleAssignmentRemoveCommand_instances, "m", _SpoFolderRoleAssignmentRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFolderRoleAssignmentRemoveCommand_instances, "m", _SpoFolderRoleAssignmentRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFolderRoleAssignmentRemoveCommand_instances, "m", _SpoFolderRoleAssignmentRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFolderRoleAssignmentRemoveCommand_instances, "m", _SpoFolderRoleAssignmentRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const removeRoleAssignment = async () => {
            if (this.verbose) {
                await logger.logToStderr(`Removing role assignment from folder in site at ${args.options.webUrl}...`);
            }
            const serverRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.folderUrl);
            const requestUrl = `${args.options.webUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}')/ListItemAllFields`;
            try {
                let principalId = args.options.principalId;
                if (args.options.upn) {
                    principalId = await this.getUserPrincipalId(args.options, logger);
                }
                else if (args.options.groupName) {
                    principalId = await this.getGroupPrincipalId(args.options, logger);
                }
                else if (args.options.entraGroupId || args.options.entraGroupName) {
                    if (this.verbose) {
                        await logger.logToStderr('Retrieving group information...');
                    }
                    const group = args.options.entraGroupId
                        ? await entraGroup.getGroupById(args.options.entraGroupId)
                        : await entraGroup.getGroupByDisplayName(args.options.entraGroupName);
                    const siteUser = await spo.ensureEntraGroup(args.options.webUrl, group);
                    principalId = siteUser.Id;
                }
                await this.removeRoleAssignment(requestUrl, principalId);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeRoleAssignment();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove a role assignment from the folder with url '${args.options.folderUrl}'?` });
            if (result) {
                await removeRoleAssignment();
            }
        }
    }
    async removeRoleAssignment(requestUrl, principalId) {
        const requestOptions = {
            url: `${requestUrl}/roleassignments/removeroleassignment(principalid='${principalId}')`,
            method: 'POST',
            headers: {
                'accept': 'application/json;odata=nometadata',
                'content-type': 'application/json'
            },
            responseType: 'json'
        };
        return request.post(requestOptions);
    }
    async getGroupPrincipalId(options, logger) {
        const group = await spo.getGroupByName(options.webUrl, options.groupName, logger, this.verbose);
        return group.Id;
    }
    async getUserPrincipalId(options, logger) {
        const user = await spo.getUserByEmail(options.webUrl, options.upn, logger, this.verbose);
        return user.Id;
    }
}
_SpoFolderRoleAssignmentRemoveCommand_instances = new WeakSet(), _SpoFolderRoleAssignmentRemoveCommand_initTelemetry = function _SpoFolderRoleAssignmentRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            principalId: typeof args.options.principalId !== 'undefined',
            upn: typeof args.options.upn !== 'undefined',
            groupName: typeof args.options.groupName !== 'undefined',
            entraGroupId: typeof args.options.entraGroupId !== 'undefined',
            entraGroupName: typeof args.options.entraGroupName !== 'undefined',
            force: !!args.options.force
        });
    });
}, _SpoFolderRoleAssignmentRemoveCommand_initOptions = function _SpoFolderRoleAssignmentRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--folderUrl <folderUrl>'
    }, {
        option: '--principalId [principalId]'
    }, {
        option: '--upn [upn]'
    }, {
        option: '--groupName [groupName]'
    }, {
        option: '--entraGroupId [entraGroupId]'
    }, {
        option: '--entraGroupName [entraGroupName]'
    }, {
        option: '-f, --force'
    });
}, _SpoFolderRoleAssignmentRemoveCommand_initValidators = function _SpoFolderRoleAssignmentRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.principalId && isNaN(args.options.principalId)) {
            return `Specified principalId ${args.options.principalId} is not a number`;
        }
        if (args.options.entraGroupId && !validation.isValidGuid(args.options.entraGroupId)) {
            return `'${args.options.entraGroupId}' is not a valid GUID for option entraGroupId.`;
        }
        const principalOptions = [args.options.principalId, args.options.upn, args.options.groupName, args.options.entraGroupId, args.options.entraGroupName];
        if (principalOptions.some(item => item !== undefined) && principalOptions.filter(item => item !== undefined).length > 1) {
            return `Specify either principalId id, upn, groupName, entraGroupId or entraGroupName`;
        }
        if (principalOptions.filter(item => item !== undefined).length === 0) {
            return `Specify at least principalId id, upn, groupName, entraGroupId or entraGroupName`;
        }
        return true;
    });
}, _SpoFolderRoleAssignmentRemoveCommand_initTypes = function _SpoFolderRoleAssignmentRemoveCommand_initTypes() {
    this.types.string.push('webUrl', 'folderUrl', 'upn', 'groupName');
    this.types.boolean.push('force');
};
export default new SpoFolderRoleAssignmentRemoveCommand();
//# sourceMappingURL=folder-roleassignment-remove.js.map