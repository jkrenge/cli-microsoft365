var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFileRoleAssignmentAddCommand_instances, _SpoFileRoleAssignmentAddCommand_initTelemetry, _SpoFileRoleAssignmentAddCommand_initOptions, _SpoFileRoleAssignmentAddCommand_initValidators, _SpoFileRoleAssignmentAddCommand_initOptionSets, _SpoFileRoleAssignmentAddCommand_initTypes;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
class SpoFileRoleAssignmentAddCommand extends SpoCommand {
    get name() {
        return commands.FILE_ROLEASSIGNMENT_ADD;
    }
    get description() {
        return 'Adds a role assignment to the specified file.';
    }
    constructor() {
        super();
        _SpoFileRoleAssignmentAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFileRoleAssignmentAddCommand_instances, "m", _SpoFileRoleAssignmentAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFileRoleAssignmentAddCommand_instances, "m", _SpoFileRoleAssignmentAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFileRoleAssignmentAddCommand_instances, "m", _SpoFileRoleAssignmentAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFileRoleAssignmentAddCommand_instances, "m", _SpoFileRoleAssignmentAddCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFileRoleAssignmentAddCommand_instances, "m", _SpoFileRoleAssignmentAddCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Adding role assignment to file in site at ${args.options.webUrl}...`);
        }
        try {
            const fileUrl = await this.getFileURL(args, logger);
            const roleDefinitionId = await this.getRoleDefinitionId(args.options, logger);
            if (args.options.upn) {
                const upnPrincipalId = await this.getUserPrincipalId(args.options, logger);
                await this.addRoleAssignment(fileUrl, args.options.webUrl, upnPrincipalId, roleDefinitionId);
            }
            else if (args.options.groupName) {
                const groupPrincipalId = await this.getGroupPrincipalId(args.options, logger);
                await this.addRoleAssignment(fileUrl, args.options.webUrl, groupPrincipalId, roleDefinitionId);
            }
            else if (args.options.entraGroupId || args.options.entraGroupName) {
                if (this.verbose) {
                    await logger.logToStderr('Retrieving group information...');
                }
                let group;
                if (args.options.entraGroupId) {
                    group = await entraGroup.getGroupById(args.options.entraGroupId);
                }
                else {
                    group = await entraGroup.getGroupByDisplayName(args.options.entraGroupName);
                }
                const entraSiteUser = await spo.ensureEntraGroup(args.options.webUrl, group);
                await this.addRoleAssignment(fileUrl, args.options.webUrl, entraSiteUser.Id, roleDefinitionId);
            }
            else {
                await this.addRoleAssignment(fileUrl, args.options.webUrl, args.options.principalId, roleDefinitionId);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async addRoleAssignment(fileUrl, webUrl, principalId, roleDefinitionId) {
        const requestOptions = {
            url: `${webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(fileUrl)}')/ListItemAllFields/roleassignments/addroleassignment(principalid='${principalId}',roledefid='${roleDefinitionId}')`,
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
        const roleDefinition = await spo.getRoleDefinitionByName(options.webUrl, options.roleDefinitionName, logger, this.verbose);
        return roleDefinition.Id;
    }
    async getGroupPrincipalId(options, logger) {
        const group = await spo.getGroupByName(options.webUrl, options.groupName, logger, this.verbose);
        return group.Id;
    }
    async getUserPrincipalId(options, logger) {
        const user = await spo.getUserByEmail(options.webUrl, options.upn, logger, this.verbose);
        return user.Id;
    }
    async getFileURL(args, logger) {
        if (args.options.fileUrl) {
            return urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl);
        }
        const file = await spo.getFileById(args.options.webUrl, args.options.fileId, logger, this.verbose);
        return file.ServerRelativeUrl;
    }
}
_SpoFileRoleAssignmentAddCommand_instances = new WeakSet(), _SpoFileRoleAssignmentAddCommand_initTelemetry = function _SpoFileRoleAssignmentAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            fileUrl: typeof args.options.fileUrl !== 'undefined',
            fileId: typeof args.options.fileId !== 'undefined',
            principalId: typeof args.options.principalId !== 'undefined',
            upn: typeof args.options.upn !== 'undefined',
            groupName: typeof args.options.groupName !== 'undefined',
            entraGroupId: typeof args.options.entraGroupId !== 'undefined',
            entraGroupName: typeof args.options.entraGroupName !== 'undefined',
            roleDefinitionId: typeof args.options.roleDefinitionId !== 'undefined',
            roleDefinitionName: typeof args.options.roleDefinitionName !== 'undefined'
        });
    });
}, _SpoFileRoleAssignmentAddCommand_initOptions = function _SpoFileRoleAssignmentAddCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--fileUrl [fileUrl]'
    }, {
        option: '-i, --fileId [fileId]'
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
}, _SpoFileRoleAssignmentAddCommand_initValidators = function _SpoFileRoleAssignmentAddCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.fileId && !validation.isValidGuid(args.options.fileId)) {
            return `${args.options.fileId} is not a valid GUID`;
        }
        if (args.options.principalId && isNaN(args.options.principalId)) {
            return `Specified principalId ${args.options.principalId} is not a number`;
        }
        if (args.options.entraGroupId && !validation.isValidGuid(args.options.entraGroupId)) {
            return `'${args.options.entraGroupId}' is not a valid GUID for option entraGroupId`;
        }
        if (args.options.roleDefinitionId && isNaN(args.options.roleDefinitionId)) {
            return `Specified roleDefinitionId ${args.options.roleDefinitionId} is not a number`;
        }
        return true;
    });
}, _SpoFileRoleAssignmentAddCommand_initOptionSets = function _SpoFileRoleAssignmentAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['fileId', 'fileUrl'] }, { options: ['principalId', 'upn', 'groupName', 'entraGroupId', 'entraGroupName'] }, { options: ['roleDefinitionId', 'roleDefinitionName'] });
}, _SpoFileRoleAssignmentAddCommand_initTypes = function _SpoFileRoleAssignmentAddCommand_initTypes() {
    this.types.string.push('webUrl', 'fileUrl', 'fileId', 'upn', 'groupName', 'entraGroupId', 'entraGroupName', 'roleDefinitionName');
};
export default new SpoFileRoleAssignmentAddCommand();
//# sourceMappingURL=file-roleassignment-add.js.map