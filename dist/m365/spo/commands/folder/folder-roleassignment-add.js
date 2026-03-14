var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFolderRoleAssignmentAddCommand_instances, _SpoFolderRoleAssignmentAddCommand_initTelemetry, _SpoFolderRoleAssignmentAddCommand_initOptions, _SpoFolderRoleAssignmentAddCommand_initValidators, _SpoFolderRoleAssignmentAddCommand_initTypes;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { spo } from '../../../../utils/spo.js';
class SpoFolderRoleAssignmentAddCommand extends SpoCommand {
    get name() {
        return commands.FOLDER_ROLEASSIGNMENT_ADD;
    }
    get description() {
        return 'Adds a role assignment to the specified folder.';
    }
    constructor() {
        super();
        _SpoFolderRoleAssignmentAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFolderRoleAssignmentAddCommand_instances, "m", _SpoFolderRoleAssignmentAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFolderRoleAssignmentAddCommand_instances, "m", _SpoFolderRoleAssignmentAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFolderRoleAssignmentAddCommand_instances, "m", _SpoFolderRoleAssignmentAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFolderRoleAssignmentAddCommand_instances, "m", _SpoFolderRoleAssignmentAddCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Adding role assignment to folder in site at ${args.options.webUrl}...`);
        }
        const serverRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.folderUrl);
        const roleFolderUrl = urlUtil.getWebRelativePath(args.options.webUrl, args.options.folderUrl);
        try {
            let requestUrl = `${args.options.webUrl}/_api/web/`;
            if (roleFolderUrl.split('/').length === 2) {
                requestUrl += `GetList('${formatting.encodeQueryParameter(serverRelativeUrl)}')`;
            }
            else {
                requestUrl += `GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}')/ListItemAllFields`;
            }
            const roleDefinitionId = await this.getRoleDefinitionId(args.options, logger);
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
            await this.addRoleAssignment(requestUrl, principalId, roleDefinitionId);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async addRoleAssignment(requestUrl, principalId, roleDefinitionId) {
        const requestOptions = {
            url: `${requestUrl}/roleassignments/addroleassignment(principalid='${principalId}',roledefid='${roleDefinitionId}')`,
            method: 'POST',
            headers: {
                'accept': 'application/json;odata=nometadata',
                'content-type': 'application/json'
            },
            responseType: 'json'
        };
        return request.post(requestOptions);
    }
    async getRoleDefinitionId(options, logger) {
        if (!options.roleDefinitionName) {
            return options.roleDefinitionId;
        }
        const roleDefintion = await spo.getRoleDefintionByName(options.webUrl, options.roleDefinitionName, logger, this.verbose);
        return roleDefintion.Id;
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
_SpoFolderRoleAssignmentAddCommand_instances = new WeakSet(), _SpoFolderRoleAssignmentAddCommand_initTelemetry = function _SpoFolderRoleAssignmentAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            principalId: typeof args.options.principalId !== 'undefined',
            upn: typeof args.options.upn !== 'undefined',
            groupName: typeof args.options.groupName !== 'undefined',
            entraGroupId: typeof args.options.entraGroupId !== 'undefined',
            entraGroupName: typeof args.options.entraGroupName !== 'undefined',
            roleDefinitionId: typeof args.options.roleDefinitionId !== 'undefined',
            roleDefinitionName: typeof args.options.roleDefinitionName !== 'undefined'
        });
    });
}, _SpoFolderRoleAssignmentAddCommand_initOptions = function _SpoFolderRoleAssignmentAddCommand_initOptions() {
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
        option: '--roleDefinitionId [roleDefinitionId]'
    }, {
        option: '--roleDefinitionName [roleDefinitionName]'
    });
}, _SpoFolderRoleAssignmentAddCommand_initValidators = function _SpoFolderRoleAssignmentAddCommand_initValidators() {
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
        if (args.options.roleDefinitionId && isNaN(args.options.roleDefinitionId)) {
            return `Specified roleDefinitionId ${args.options.roleDefinitionId} is not a number`;
        }
        const principalOptions = [args.options.principalId, args.options.upn, args.options.groupName, args.options.entraGroupId, args.options.entraGroupName];
        if (!principalOptions.some(item => item !== undefined)) {
            return `Specify either principalId, upn, groupName, entraGroupId or entraGroupName`;
        }
        if (principalOptions.filter(item => item !== undefined).length > 1) {
            return `Specify either principalId, upn, groupName, entraGroupId or entraGroupName but not multiple`;
        }
        const roleDefinitionOptions = [args.options.roleDefinitionId, args.options.roleDefinitionName];
        if (!roleDefinitionOptions.some(item => item !== undefined)) {
            return `Specify either roleDefinitionId id or roleDefinitionName`;
        }
        if (roleDefinitionOptions.filter(item => item !== undefined).length > 1) {
            return `Specify either roleDefinitionId id or roleDefinitionName but not both`;
        }
        return true;
    });
}, _SpoFolderRoleAssignmentAddCommand_initTypes = function _SpoFolderRoleAssignmentAddCommand_initTypes() {
    this.types.string.push('webUrl', 'folderUrl', 'upn', 'groupName', 'roleDefinitionName');
};
export default new SpoFolderRoleAssignmentAddCommand();
//# sourceMappingURL=folder-roleassignment-add.js.map