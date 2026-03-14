var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListItemRoleAssignmentRemoveCommand_instances, _SpoListItemRoleAssignmentRemoveCommand_initTelemetry, _SpoListItemRoleAssignmentRemoveCommand_initOptions, _SpoListItemRoleAssignmentRemoveCommand_initValidators, _SpoListItemRoleAssignmentRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { spo } from '../../../../utils/spo.js';
class SpoListItemRoleAssignmentRemoveCommand extends SpoCommand {
    get name() {
        return commands.LISTITEM_ROLEASSIGNMENT_REMOVE;
    }
    get description() {
        return 'Removes a role assignment from list item permissions';
    }
    constructor() {
        super();
        _SpoListItemRoleAssignmentRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListItemRoleAssignmentRemoveCommand_instances, "m", _SpoListItemRoleAssignmentRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListItemRoleAssignmentRemoveCommand_instances, "m", _SpoListItemRoleAssignmentRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListItemRoleAssignmentRemoveCommand_instances, "m", _SpoListItemRoleAssignmentRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListItemRoleAssignmentRemoveCommand_instances, "m", _SpoListItemRoleAssignmentRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeRoleAssignment(logger, args.options);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove role assignment from listitem ${args.options.listItemId} from list ${args.options.listId || args.options.listTitle} from site ${args.options.webUrl}?` });
            if (result) {
                await this.removeRoleAssignment(logger, args.options);
            }
        }
    }
    async removeRoleAssignment(logger, options) {
        if (this.verbose) {
            await logger.logToStderr(`Removing role assignment from listitem in site at ${options.webUrl}...`);
        }
        try {
            let requestUrl = `${options.webUrl}/_api/web/`;
            if (options.listId) {
                requestUrl += `lists(guid'${formatting.encodeQueryParameter(options.listId)}')/`;
            }
            else if (options.listTitle) {
                requestUrl += `lists/getByTitle('${formatting.encodeQueryParameter(options.listTitle)}')/`;
            }
            else if (options.listUrl) {
                const listServerRelativeUrl = urlUtil.getServerRelativePath(options.webUrl, options.listUrl);
                requestUrl += `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/`;
            }
            requestUrl += `items(${options.listItemId})/`;
            if (options.upn) {
                options.principalId = await this.getUserPrincipalId(options, logger);
            }
            else if (options.groupName) {
                options.principalId = await this.getGroupPrincipalId(options, logger);
            }
            else if (options.entraGroupId || options.entraGroupName) {
                if (this.verbose) {
                    await logger.logToStderr('Retrieving group information...');
                }
                const group = options.entraGroupId
                    ? await entraGroup.getGroupById(options.entraGroupId)
                    : await entraGroup.getGroupByDisplayName(options.entraGroupName);
                const siteUser = await spo.ensureEntraGroup(options.webUrl, group);
                options.principalId = siteUser.Id;
            }
            await this.removeRoleAssignmentWithRequestUrl(requestUrl, logger, options);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async removeRoleAssignmentWithRequestUrl(requestUrl, logger, options) {
        const requestOptions = {
            url: `${requestUrl}roleassignments/removeroleassignment(principalid='${options.principalId}')`,
            method: 'POST',
            headers: {
                'accept': 'application/json;odata=nometadata',
                'content-type': 'application/json'
            },
            responseType: 'json'
        };
        await request.post(requestOptions);
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
_SpoListItemRoleAssignmentRemoveCommand_instances = new WeakSet(), _SpoListItemRoleAssignmentRemoveCommand_initTelemetry = function _SpoListItemRoleAssignmentRemoveCommand_initTelemetry() {
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
            force: (!(!args.options.force)).toString()
        });
    });
}, _SpoListItemRoleAssignmentRemoveCommand_initOptions = function _SpoListItemRoleAssignmentRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-i, --listId [listId]'
    }, {
        option: '-t, --listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '--listItemId <listItemId>'
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
}, _SpoListItemRoleAssignmentRemoveCommand_initValidators = function _SpoListItemRoleAssignmentRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} is not a valid GUID`;
        }
        if (args.options.listItemId && isNaN(args.options.listItemId)) {
            return `Specified listItemId ${args.options.listItemId} is not a number`;
        }
        if (args.options.principalId && isNaN(args.options.principalId)) {
            return `Specified principalId ${args.options.principalId} is not a number`;
        }
        if (args.options.entraGroupId && !validation.isValidGuid(args.options.entraGroupId)) {
            return `'${args.options.entraGroupId}' is not a valid GUID for option entraGroupId.`;
        }
        return true;
    });
}, _SpoListItemRoleAssignmentRemoveCommand_initOptionSets = function _SpoListItemRoleAssignmentRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] }, { options: ['principalId', 'upn', 'groupName', 'entraGroupId', 'entraGroupName'] });
};
export default new SpoListItemRoleAssignmentRemoveCommand();
//# sourceMappingURL=listitem-roleassignment-remove.js.map