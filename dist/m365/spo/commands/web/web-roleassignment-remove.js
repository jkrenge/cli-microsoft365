var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoWebRoleAssignmentRemoveCommand_instances, _SpoWebRoleAssignmentRemoveCommand_initTelemetry, _SpoWebRoleAssignmentRemoveCommand_initOptions, _SpoWebRoleAssignmentRemoveCommand_initValidators, _SpoWebRoleAssignmentRemoveCommand_initOptionSets;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { spo } from '../../../../utils/spo.js';
import { cli } from '../../../../cli/cli.js';
class SpoWebRoleAssignmentRemoveCommand extends SpoCommand {
    get name() {
        return commands.WEB_ROLEASSIGNMENT_REMOVE;
    }
    get description() {
        return 'Removes a role assignment from web permissions';
    }
    constructor() {
        super();
        _SpoWebRoleAssignmentRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoWebRoleAssignmentRemoveCommand_instances, "m", _SpoWebRoleAssignmentRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoWebRoleAssignmentRemoveCommand_instances, "m", _SpoWebRoleAssignmentRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoWebRoleAssignmentRemoveCommand_instances, "m", _SpoWebRoleAssignmentRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoWebRoleAssignmentRemoveCommand_instances, "m", _SpoWebRoleAssignmentRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeRoleAssignment(logger, args.options);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove role assignment from web ${args.options.webUrl}?` });
            if (result) {
                await this.removeRoleAssignment(logger, args.options);
            }
        }
    }
    async removeRoleAssignment(logger, options) {
        if (this.verbose) {
            await logger.logToStderr(`Removing role assignment from web ${options.webUrl}...`);
        }
        try {
            if (options.upn) {
                const principalId = await this.getUserPrincipalId(options, logger);
                await this.removeRoleAssignmentWithOptions(options.webUrl, principalId, logger);
            }
            else if (options.groupName) {
                const principalId = await this.getGroupPrincipalId(options, logger);
                await this.removeRoleAssignmentWithOptions(options.webUrl, principalId, logger);
            }
            else if (options.entraGroupId || options.entraGroupName) {
                if (this.verbose) {
                    await logger.logToStderr('Retrieving group information...');
                }
                const group = options.entraGroupId
                    ? await entraGroup.getGroupById(options.entraGroupId)
                    : await entraGroup.getGroupByDisplayName(options.entraGroupName);
                const siteUser = await spo.ensureEntraGroup(options.webUrl, group);
                await this.removeRoleAssignmentWithOptions(options.webUrl, siteUser.Id, logger);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async removeRoleAssignmentWithOptions(webUrl, principalId, logger) {
        if (this.verbose) {
            await logger.logToStderr('Removing role assignment...');
        }
        const requestOptions = {
            url: `${webUrl}/_api/web/roleassignments/removeroleassignment(principalid='${principalId}')`,
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
_SpoWebRoleAssignmentRemoveCommand_instances = new WeakSet(), _SpoWebRoleAssignmentRemoveCommand_initTelemetry = function _SpoWebRoleAssignmentRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            principalId: typeof args.options.principalId !== 'undefined',
            upn: typeof args.options.upn !== 'undefined',
            groupName: typeof args.options.groupName !== 'undefined',
            entraGroupId: typeof args.options.entraGroupId !== 'undefined',
            entraGroupName: typeof args.options.entraGroupName !== 'undefined',
            force: (!(!args.options.force)).toString()
        });
    });
}, _SpoWebRoleAssignmentRemoveCommand_initOptions = function _SpoWebRoleAssignmentRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
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
}, _SpoWebRoleAssignmentRemoveCommand_initValidators = function _SpoWebRoleAssignmentRemoveCommand_initValidators() {
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
        return true;
    });
}, _SpoWebRoleAssignmentRemoveCommand_initOptionSets = function _SpoWebRoleAssignmentRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['principalId', 'upn', 'groupName', 'entraGroupId', 'entraGroupName'] });
};
export default new SpoWebRoleAssignmentRemoveCommand();
//# sourceMappingURL=web-roleassignment-remove.js.map