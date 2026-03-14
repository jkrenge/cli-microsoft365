var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _FlowOwnerEnsureCommand_instances, _a, _FlowOwnerEnsureCommand_initTelemetry, _FlowOwnerEnsureCommand_initOptions, _FlowOwnerEnsureCommand_initOptionSets, _FlowOwnerEnsureCommand_initValidators;
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import PowerAutomateCommand from '../../../base/PowerAutomateCommand.js';
import commands from '../../commands.js';
class FlowOwnerEnsureCommand extends PowerAutomateCommand {
    get name() {
        return commands.OWNER_ENSURE;
    }
    get description() {
        return 'Assigns/updates permissions to a Power Automate flow';
    }
    constructor() {
        super();
        _FlowOwnerEnsureCommand_instances.add(this);
        __classPrivateFieldGet(this, _FlowOwnerEnsureCommand_instances, "m", _FlowOwnerEnsureCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _FlowOwnerEnsureCommand_instances, "m", _FlowOwnerEnsureCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _FlowOwnerEnsureCommand_instances, "m", _FlowOwnerEnsureCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _FlowOwnerEnsureCommand_instances, "m", _FlowOwnerEnsureCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Assigning permissions for ${args.options.userId || args.options.userName || args.options.groupId || args.options.groupName} with permissions ${args.options.roleName} to Power Automate flow ${args.options.flowName}`);
            }
            let id = '';
            if (args.options.userId) {
                id = args.options.userId;
            }
            else if (args.options.userName) {
                id = await entraUser.getUserIdByUpn(args.options.userName);
            }
            else if (args.options.groupId) {
                id = args.options.groupId;
            }
            else {
                id = await entraGroup.getGroupIdByDisplayName(args.options.groupName);
            }
            let type;
            if (args.options.userId || args.options.userName) {
                type = 'User';
            }
            else {
                type = 'Group';
            }
            const requestOptions = {
                url: `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/${args.options.asAdmin ? 'scopes/admin/' : ''}environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.flowName)}/modifyPermissions?api-version=2016-11-01`,
                headers: {
                    accept: 'application/json'
                },
                data: {
                    put: [
                        {
                            properties: {
                                principal: {
                                    id: id,
                                    type: type
                                },
                                roleName: args.options.roleName
                            }
                        }
                    ]
                },
                responseType: 'json'
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_a = FlowOwnerEnsureCommand, _FlowOwnerEnsureCommand_instances = new WeakSet(), _FlowOwnerEnsureCommand_initTelemetry = function _FlowOwnerEnsureCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            asAdmin: !!args.options.asAdmin,
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            groupId: typeof args.options.groupId !== 'undefined',
            groupName: typeof args.options.groupName !== 'undefined'
        });
    });
}, _FlowOwnerEnsureCommand_initOptions = function _FlowOwnerEnsureCommand_initOptions() {
    this.options.unshift({
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '--flowName <flowName>'
    }, {
        option: '--userId [userId]'
    }, {
        option: '--userName [userName]'
    }, {
        option: '--groupId [groupId]'
    }, {
        option: '--groupName [groupName]'
    }, {
        option: '--roleName <roleName>',
        autocomplete: _a.allowedRoleNames
    }, {
        option: '--asAdmin'
    });
}, _FlowOwnerEnsureCommand_initOptionSets = function _FlowOwnerEnsureCommand_initOptionSets() {
    this.optionSets.push({ options: ['userId', 'userName', 'groupId', 'groupName'] });
}, _FlowOwnerEnsureCommand_initValidators = function _FlowOwnerEnsureCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.flowName)) {
            return `${args.options.flowName} is not a valid GUID.`;
        }
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID.`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid userName.`;
        }
        if (args.options.groupId && !validation.isValidGuid(args.options.groupId)) {
            return `${args.options.groupId} is not a valid GUID.`;
        }
        if (_a.allowedRoleNames.indexOf(args.options.roleName) === -1) {
            return `${args.options.roleName} is not a valid roleName. Valid values are: ${_a.allowedRoleNames.join(', ')}`;
        }
        return true;
    });
};
FlowOwnerEnsureCommand.allowedRoleNames = ['CanView', 'CanEdit'];
export default new FlowOwnerEnsureCommand();
//# sourceMappingURL=owner-ensure.js.map