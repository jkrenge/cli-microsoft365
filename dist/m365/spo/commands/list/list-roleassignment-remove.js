var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListRoleAssignmentRemoveCommand_instances, _SpoListRoleAssignmentRemoveCommand_initTelemetry, _SpoListRoleAssignmentRemoveCommand_initOptions, _SpoListRoleAssignmentRemoveCommand_initValidators, _SpoListRoleAssignmentRemoveCommand_initOptionSets, _SpoListRoleAssignmentRemoveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
class SpoListRoleAssignmentRemoveCommand extends SpoCommand {
    get name() {
        return commands.LIST_ROLEASSIGNMENT_REMOVE;
    }
    get description() {
        return 'Removes a role assignment from list permissions';
    }
    constructor() {
        super();
        _SpoListRoleAssignmentRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListRoleAssignmentRemoveCommand_instances, "m", _SpoListRoleAssignmentRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListRoleAssignmentRemoveCommand_instances, "m", _SpoListRoleAssignmentRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListRoleAssignmentRemoveCommand_instances, "m", _SpoListRoleAssignmentRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListRoleAssignmentRemoveCommand_instances, "m", _SpoListRoleAssignmentRemoveCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoListRoleAssignmentRemoveCommand_instances, "m", _SpoListRoleAssignmentRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const removeRoleAssignment = async () => {
            if (this.verbose) {
                await logger.logToStderr(`Removing role assignment from list '${args.options.listId || args.options.listTitle || args.options.listUrl}' of site ${args.options.webUrl}...`);
            }
            try {
                let requestUrl = `${args.options.webUrl}/_api/web/`;
                if (args.options.listId) {
                    requestUrl += `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/`;
                }
                else if (args.options.listTitle) {
                    requestUrl += `lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')/`;
                }
                else if (args.options.listUrl) {
                    const listServerRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
                    requestUrl += `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/`;
                }
                let principalId = args.options.principalId;
                if (args.options.upn) {
                    const user = await spo.ensureUser(args.options.webUrl, args.options.upn);
                    principalId = user.Id;
                }
                else if (args.options.groupName) {
                    const spGroup = await spo.getGroupByName(args.options.webUrl, args.options.groupName, logger, this.verbose);
                    principalId = spGroup.Id;
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
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove role assignment from the specified user of list '${args.options.listId || args.options.listTitle || args.options.listUrl}'?` });
            if (result) {
                await removeRoleAssignment();
            }
        }
    }
    async removeRoleAssignment(requestUrl, principalId) {
        const requestOptions = {
            url: `${requestUrl}roleassignments/removeroleassignment(principalid='${principalId}')`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        return request.post(requestOptions);
    }
}
_SpoListRoleAssignmentRemoveCommand_instances = new WeakSet(), _SpoListRoleAssignmentRemoveCommand_initTelemetry = function _SpoListRoleAssignmentRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            principalId: typeof args.options.principalId !== 'undefined',
            upn: typeof args.options.upn !== 'undefined',
            groupName: typeof args.options.groupName !== 'undefined',
            entraGroupId: typeof args.options.entraGroupId !== 'undefined',
            entraGroupName: typeof args.options.entraGroupName !== 'undefined',
            force: !!args.options.force
        });
    });
}, _SpoListRoleAssignmentRemoveCommand_initOptions = function _SpoListRoleAssignmentRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-i, --listId [listId]'
    }, {
        option: '-t, --listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
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
}, _SpoListRoleAssignmentRemoveCommand_initValidators = function _SpoListRoleAssignmentRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
            return `'${args.options.listId}' is not a valid GUID for option listId.`;
        }
        if (args.options.upn && !validation.isValidUserPrincipalName(args.options.upn)) {
            return `'${args.options.upn}' is not a valid user principal name for option upn.`;
        }
        if (args.options.principalId && !validation.isValidPositiveInteger(args.options.principalId)) {
            return `'${args.options.principalId}' is not a valid number for option principalId.`;
        }
        if (args.options.entraGroupId && !validation.isValidGuid(args.options.entraGroupId)) {
            return `'${args.options.entraGroupId}' is not a valid GUID for option entraGroupId.`;
        }
        return true;
    });
}, _SpoListRoleAssignmentRemoveCommand_initOptionSets = function _SpoListRoleAssignmentRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] }, { options: ['principalId', 'upn', 'groupName', 'entraGroupId', 'entraGroupName'] });
}, _SpoListRoleAssignmentRemoveCommand_initTypes = function _SpoListRoleAssignmentRemoveCommand_initTypes() {
    this.types.string.push('webUrl', 'listId', 'listTitle', 'listUrl', 'upn', 'groupName', 'entraGroupId', 'entraGroupName');
    this.types.boolean.push('force');
};
export default new SpoListRoleAssignmentRemoveCommand();
//# sourceMappingURL=list-roleassignment-remove.js.map