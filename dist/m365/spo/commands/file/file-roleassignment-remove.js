var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFileRoleAssignmentRemoveCommand_instances, _SpoFileRoleAssignmentRemoveCommand_initTelemetry, _SpoFileRoleAssignmentRemoveCommand_initOptions, _SpoFileRoleAssignmentRemoveCommand_initValidators, _SpoFileRoleAssignmentRemoveCommand_initOptionSets, _SpoFileRoleAssignmentRemoveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
class SpoFileRoleAssignmentRemoveCommand extends SpoCommand {
    get name() {
        return commands.FILE_ROLEASSIGNMENT_REMOVE;
    }
    get description() {
        return 'Removes a role assignment from a file.';
    }
    constructor() {
        super();
        _SpoFileRoleAssignmentRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFileRoleAssignmentRemoveCommand_instances, "m", _SpoFileRoleAssignmentRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFileRoleAssignmentRemoveCommand_instances, "m", _SpoFileRoleAssignmentRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFileRoleAssignmentRemoveCommand_instances, "m", _SpoFileRoleAssignmentRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFileRoleAssignmentRemoveCommand_instances, "m", _SpoFileRoleAssignmentRemoveCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFileRoleAssignmentRemoveCommand_instances, "m", _SpoFileRoleAssignmentRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const removeRoleAssignment = async () => {
            if (this.verbose) {
                await logger.logToStderr(`Removing role assignment for ${args.options.groupName || args.options.upn} from file ${args.options.fileUrl || args.options.fileId}`);
            }
            try {
                const fileURL = await this.getFileURL(args, logger);
                let principalId;
                if (args.options.groupName) {
                    principalId = await this.getGroupPrincipalId(args.options, logger);
                }
                else if (args.options.upn) {
                    principalId = await this.getUserPrincipalId(args.options, logger);
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
                    principalId = entraSiteUser.Id;
                }
                else {
                    principalId = args.options.principalId;
                }
                const serverRelativePath = urlUtil.getServerRelativePath(args.options.webUrl, fileURL);
                const requestOptions = {
                    url: `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')/ListItemAllFields/roleassignments/removeroleassignment(principalid='${principalId}')`,
                    headers: {
                        accept: 'application/json;odata=nometadata'
                    },
                    responseType: 'json'
                };
                await request.post(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeRoleAssignment();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove role assignment from file ${args.options.fileUrl || args.options.fileId} from site ${args.options.webUrl}?` });
            if (result) {
                await removeRoleAssignment();
            }
        }
    }
    async getFileURL(args, logger) {
        if (args.options.fileUrl) {
            return urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl);
        }
        const file = await spo.getFileById(args.options.webUrl, args.options.fileId, logger, this.verbose);
        return file.ServerRelativeUrl;
    }
    async getUserPrincipalId(options, logger) {
        const user = await spo.getUserByEmail(options.webUrl, options.upn, logger, this.verbose);
        return user.Id;
    }
    async getGroupPrincipalId(options, logger) {
        const group = await spo.getGroupByName(options.webUrl, options.groupName, logger, this.verbose);
        return group.Id;
    }
}
_SpoFileRoleAssignmentRemoveCommand_instances = new WeakSet(), _SpoFileRoleAssignmentRemoveCommand_initTelemetry = function _SpoFileRoleAssignmentRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            fileUrl: typeof args.options.fileUrl !== 'undefined',
            fileId: typeof args.options.fileId !== 'undefined',
            principalId: typeof args.options.principalId !== 'undefined',
            upn: typeof args.options.upn !== 'undefined',
            groupName: typeof args.options.groupName !== 'undefined',
            entraGroupId: typeof args.options.entraGroupId !== 'undefined',
            entraGroupName: typeof args.options.entraGroupName !== 'undefined',
            force: !!args.options.force
        });
    });
}, _SpoFileRoleAssignmentRemoveCommand_initOptions = function _SpoFileRoleAssignmentRemoveCommand_initOptions() {
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
        option: '-f, --force'
    });
}, _SpoFileRoleAssignmentRemoveCommand_initValidators = function _SpoFileRoleAssignmentRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.principalId && isNaN(args.options.principalId)) {
            return `Specified principalId ${args.options.principalId} is not a number`;
        }
        if (args.options.entraGroupId && !validation.isValidGuid(args.options.entraGroupId)) {
            return `'${args.options.entraGroupId}' is not a valid GUID for option entraGroupId`;
        }
        if (args.options.fileId && !validation.isValidGuid(args.options.fileId)) {
            return `${args.options.fileId} is not a valid GUID`;
        }
        return true;
    });
}, _SpoFileRoleAssignmentRemoveCommand_initOptionSets = function _SpoFileRoleAssignmentRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['fileUrl', 'fileId'] }, { options: ['upn', 'groupName', 'principalId', 'entraGroupId', 'entraGroupName'] });
}, _SpoFileRoleAssignmentRemoveCommand_initTypes = function _SpoFileRoleAssignmentRemoveCommand_initTypes() {
    this.types.string.push('webUrl', 'fileUrl', 'fileId', 'upn', 'groupName', 'entraGroupId', 'entraGroupName');
    this.types.boolean.push('force');
};
export default new SpoFileRoleAssignmentRemoveCommand();
//# sourceMappingURL=file-roleassignment-remove.js.map