var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoWebRoleAssignmentAddCommand_instances, _SpoWebRoleAssignmentAddCommand_initTelemetry, _SpoWebRoleAssignmentAddCommand_initOptions, _SpoWebRoleAssignmentAddCommand_initValidators, _SpoWebRoleAssignmentAddCommand_initOptionSets;
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoWebRoleAssignmentAddCommand extends SpoCommand {
    get name() {
        return commands.WEB_ROLEASSIGNMENT_ADD;
    }
    get description() {
        return 'Adds a role assignment to web';
    }
    constructor() {
        super();
        _SpoWebRoleAssignmentAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoWebRoleAssignmentAddCommand_instances, "m", _SpoWebRoleAssignmentAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoWebRoleAssignmentAddCommand_instances, "m", _SpoWebRoleAssignmentAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoWebRoleAssignmentAddCommand_instances, "m", _SpoWebRoleAssignmentAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoWebRoleAssignmentAddCommand_instances, "m", _SpoWebRoleAssignmentAddCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Adding role assignment to web ${args.options.webUrl}...`);
        }
        try {
            const roleDefinitionId = await this.getRoleDefinitionId(args.options, logger);
            if (args.options.upn) {
                const principalId = await this.getUserPrincipalId(args.options, logger);
                await this.addRoleAssignment(args.options.webUrl, principalId, roleDefinitionId, logger);
            }
            else if (args.options.groupName) {
                const principalId = await this.getGroupPrincipalId(args.options, logger);
                await this.addRoleAssignment(args.options.webUrl, principalId, roleDefinitionId, logger);
            }
            else if (args.options.entraGroupId || args.options.entraGroupName) {
                if (this.verbose) {
                    await logger.logToStderr('Retrieving group information...');
                }
                const group = args.options.entraGroupId
                    ? await entraGroup.getGroupById(args.options.entraGroupId)
                    : await entraGroup.getGroupByDisplayName(args.options.entraGroupName);
                const siteUser = await spo.ensureEntraGroup(args.options.webUrl, group);
                await this.addRoleAssignment(args.options.webUrl, siteUser.Id, roleDefinitionId, logger);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async addRoleAssignment(webUrl, principalId, roleDefinitionId, logger) {
        if (this.verbose) {
            await logger.logToStderr('Adding role assignment...');
        }
        const requestOptions = {
            url: `${webUrl}/_api/web/roleassignments/addroleassignment(principalid='${principalId}',roledefid='${roleDefinitionId}')`,
            method: 'POST',
            headers: {
                'accept': 'application/json;odata=nometadata',
                'content-type': 'application/json'
            },
            responseType: 'json'
        };
        await request.post(requestOptions);
    }
    async getRoleDefinitionId(options, logger) {
        if (!options.roleDefinitionName) {
            return options.roleDefinitionId;
        }
        const roledefinition = await spo.getRoleDefinitionByName(options.webUrl, options.roleDefinitionName, logger, this.verbose);
        return roledefinition.Id;
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
_SpoWebRoleAssignmentAddCommand_instances = new WeakSet(), _SpoWebRoleAssignmentAddCommand_initTelemetry = function _SpoWebRoleAssignmentAddCommand_initTelemetry() {
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
}, _SpoWebRoleAssignmentAddCommand_initOptions = function _SpoWebRoleAssignmentAddCommand_initOptions() {
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
        option: '--roleDefinitionId [roleDefinitionId]'
    }, {
        option: '--roleDefinitionName [roleDefinitionName]'
    });
}, _SpoWebRoleAssignmentAddCommand_initValidators = function _SpoWebRoleAssignmentAddCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.principalId && isNaN(args.options.principalId)) {
            return `Specified principalId ${args.options.principalId} is not a number`;
        }
        if (args.options.roleDefinitionId && isNaN(args.options.roleDefinitionId)) {
            return `Specified roleDefinitionId ${args.options.roleDefinitionId} is not a number`;
        }
        if (args.options.entraGroupId && !validation.isValidGuid(args.options.entraGroupId)) {
            return `'${args.options.entraGroupId}' is not a valid GUID for option entraGroupId.`;
        }
        return true;
    });
}, _SpoWebRoleAssignmentAddCommand_initOptionSets = function _SpoWebRoleAssignmentAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['principalId', 'upn', 'groupName', 'entraGroupId', 'entraGroupName'] }, { options: ['roleDefinitionId', 'roleDefinitionName'] });
};
export default new SpoWebRoleAssignmentAddCommand();
//# sourceMappingURL=web-roleassignment-add.js.map