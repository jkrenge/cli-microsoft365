var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraAdministrativeUnitMemberRemoveCommand_instances, _EntraAdministrativeUnitMemberRemoveCommand_initTelemetry, _EntraAdministrativeUnitMemberRemoveCommand_initOptions, _EntraAdministrativeUnitMemberRemoveCommand_initValidators, _EntraAdministrativeUnitMemberRemoveCommand_initOptionSets;
import { entraAdministrativeUnit } from '../../../../utils/entraAdministrativeUnit.js';
import { entraDevice } from '../../../../utils/entraDevice.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
import { cli } from '../../../../cli/cli.js';
class EntraAdministrativeUnitMemberRemoveCommand extends GraphCommand {
    get name() {
        return commands.ADMINISTRATIVEUNIT_MEMBER_REMOVE;
    }
    get description() {
        return 'Remove a specific member (user, group, or device) from an administrative unit';
    }
    constructor() {
        super();
        _EntraAdministrativeUnitMemberRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraAdministrativeUnitMemberRemoveCommand_instances, "m", _EntraAdministrativeUnitMemberRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraAdministrativeUnitMemberRemoveCommand_instances, "m", _EntraAdministrativeUnitMemberRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraAdministrativeUnitMemberRemoveCommand_instances, "m", _EntraAdministrativeUnitMemberRemoveCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _EntraAdministrativeUnitMemberRemoveCommand_instances, "m", _EntraAdministrativeUnitMemberRemoveCommand_initTelemetry).call(this);
    }
    async commandAction(logger, args) {
        const removeAdministrativeUnitMember = async () => {
            let administrativeUnitId = args.options.administrativeUnitId;
            let memberId = args.options.id;
            try {
                if (args.options.administrativeUnitName) {
                    if (this.verbose) {
                        await logger.logToStderr(`Retrieving Administrative Unit Id...`);
                    }
                    administrativeUnitId = (await entraAdministrativeUnit.getAdministrativeUnitByDisplayName(args.options.administrativeUnitName)).id;
                }
                if (args.options.userId || args.options.userName) {
                    memberId = args.options.userId;
                    if (args.options.userName) {
                        if (this.verbose) {
                            await logger.logToStderr(`Retrieving User Id...`);
                        }
                        memberId = await entraUser.getUserIdByUpn(args.options.userName);
                    }
                }
                else if (args.options.groupId || args.options.groupName) {
                    memberId = args.options.groupId;
                    if (args.options.groupName) {
                        if (this.verbose) {
                            await logger.logToStderr(`Retrieving Group Id...`);
                        }
                        memberId = await entraGroup.getGroupIdByDisplayName(args.options.groupName);
                    }
                }
                else if (args.options.deviceId || args.options.deviceName) {
                    memberId = args.options.deviceId;
                    if (args.options.deviceName) {
                        if (this.verbose) {
                            await logger.logToStderr(`Retrieving Device Id`);
                        }
                        memberId = (await entraDevice.getDeviceByDisplayName(args.options.deviceName)).id;
                    }
                }
                const requestOptions = {
                    url: `${this.resource}/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/${memberId}/$ref`,
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    }
                };
                await request.delete(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeAdministrativeUnitMember();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove member '${args.options.id || args.options.userName || args.options.groupName || args.options.deviceName}' from administrative unit '${args.options.administrativeUnitId || args.options.administrativeUnitName}'?` });
            if (result) {
                await removeAdministrativeUnitMember();
            }
        }
    }
}
_EntraAdministrativeUnitMemberRemoveCommand_instances = new WeakSet(), _EntraAdministrativeUnitMemberRemoveCommand_initTelemetry = function _EntraAdministrativeUnitMemberRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            groupId: typeof args.options.groupId !== 'undefined',
            groupName: typeof args.options.groupName !== 'undefined',
            deviceId: typeof args.options.deviceId !== 'undefined',
            deviceName: typeof args.options.deviceName !== 'undefined',
            administrativeUnitId: typeof args.options.administrativeUnitId !== 'undefined',
            administrativeUnitName: typeof args.options.administrativeUnitName !== 'undefined',
            force: !!args.options.force
        });
    });
}, _EntraAdministrativeUnitMemberRemoveCommand_initOptions = function _EntraAdministrativeUnitMemberRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: "--userId [userId]"
    }, {
        option: "--userName [userName]"
    }, {
        option: "--groupId [groupId]"
    }, {
        option: "--groupName [groupName]"
    }, {
        option: "--deviceId [deviceId]"
    }, {
        option: "--deviceName [deviceName]"
    }, {
        option: '--administrativeUnitId [administrativeUnitId]'
    }, {
        option: '--administrativeUnitName [administrativeUnitName]'
    }, {
        option: '-f, --force'
    });
}, _EntraAdministrativeUnitMemberRemoveCommand_initValidators = function _EntraAdministrativeUnitMemberRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.administrativeUnitId && !validation.isValidGuid(args.options.administrativeUnitId)) {
            return `${args.options.administrativeUnitId} is not a valid GUID`;
        }
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID`;
        }
        if (args.options.groupId && !validation.isValidGuid(args.options.groupId)) {
            return `${args.options.groupId} is not a valid GUID`;
        }
        if (args.options.deviceId && !validation.isValidGuid(args.options.deviceId)) {
            return `${args.options.deviceId} is not a valid GUID`;
        }
        return true;
    });
}, _EntraAdministrativeUnitMemberRemoveCommand_initOptionSets = function _EntraAdministrativeUnitMemberRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['administrativeUnitId', 'administrativeUnitName'] });
    this.optionSets.push({ options: ['id', 'userId', 'userName', 'groupId', 'groupName', 'deviceId', 'deviceName'] });
};
export default new EntraAdministrativeUnitMemberRemoveCommand();
//# sourceMappingURL=administrativeunit-member-remove.js.map