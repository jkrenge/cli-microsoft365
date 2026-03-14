var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _FlowOwnerRemoveCommand_instances, _FlowOwnerRemoveCommand_initTelemetry, _FlowOwnerRemoveCommand_initOptions, _FlowOwnerRemoveCommand_initValidators, _FlowOwnerRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import PowerAutomateCommand from '../../../base/PowerAutomateCommand.js';
import commands from '../../commands.js';
class FlowOwnerRemoveCommand extends PowerAutomateCommand {
    get name() {
        return commands.OWNER_REMOVE;
    }
    get description() {
        return 'Removes owner permissions to a Power Automate flow';
    }
    constructor() {
        super();
        _FlowOwnerRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _FlowOwnerRemoveCommand_instances, "m", _FlowOwnerRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _FlowOwnerRemoveCommand_instances, "m", _FlowOwnerRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _FlowOwnerRemoveCommand_instances, "m", _FlowOwnerRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _FlowOwnerRemoveCommand_instances, "m", _FlowOwnerRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Removing owner ${args.options.userId || args.options.userName || args.options.groupId || args.options.groupName} from flow ${args.options.flowName} in environment ${args.options.environmentName}`);
            }
            const removeFlowOwner = async () => {
                let idToRemove;
                if (args.options.userId) {
                    idToRemove = args.options.userId;
                }
                else if (args.options.userName) {
                    idToRemove = await entraUser.getUserIdByUpn(args.options.userName);
                }
                else if (args.options.groupId) {
                    idToRemove = args.options.groupId;
                }
                else {
                    idToRemove = await entraGroup.getGroupIdByDisplayName(args.options.groupName);
                }
                const requestOptions = {
                    url: `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/${args.options.asAdmin ? 'scopes/admin/' : ''}environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.flowName)}/modifyPermissions?api-version=2016-11-01`,
                    headers: {
                        accept: 'application/json'
                    },
                    data: {
                        delete: [
                            {
                                id: idToRemove
                            }
                        ]
                    },
                    responseType: 'json'
                };
                await request.post(requestOptions);
            };
            if (args.options.force) {
                await removeFlowOwner();
            }
            else {
                const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove owner '${args.options.groupId || args.options.groupName || args.options.userId || args.options.userName}' from the specified flow?` });
                if (result) {
                    await removeFlowOwner();
                }
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_FlowOwnerRemoveCommand_instances = new WeakSet(), _FlowOwnerRemoveCommand_initTelemetry = function _FlowOwnerRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            groupId: typeof args.options.groupId !== 'undefined',
            groupName: typeof args.options.groupName !== 'undefined',
            asAdmin: !!args.options.asAdmin,
            force: !!args.options.force
        });
    });
}, _FlowOwnerRemoveCommand_initOptions = function _FlowOwnerRemoveCommand_initOptions() {
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
        option: '--asAdmin'
    }, {
        option: '-f, --force'
    });
}, _FlowOwnerRemoveCommand_initValidators = function _FlowOwnerRemoveCommand_initValidators() {
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
        return true;
    });
}, _FlowOwnerRemoveCommand_initOptionSets = function _FlowOwnerRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['userId', 'userName', 'groupId', 'groupName'] });
};
export default new FlowOwnerRemoveCommand();
//# sourceMappingURL=owner-remove.js.map