var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraAdministrativeUnitMemberAddCommand_instances, _EntraAdministrativeUnitMemberAddCommand_initTelemetry, _EntraAdministrativeUnitMemberAddCommand_initOptions, _EntraAdministrativeUnitMemberAddCommand_initValidators, _EntraAdministrativeUnitMemberAddCommand_initOptionSets;
import { entraAdministrativeUnit } from "../../../../utils/entraAdministrativeUnit.js";
import { entraGroup } from "../../../../utils/entraGroup.js";
import { entraUser } from "../../../../utils/entraUser.js";
import { validation } from "../../../../utils/validation.js";
import GraphCommand from "../../../base/GraphCommand.js";
import commands from "../../commands.js";
import request from "../../../../request.js";
import { entraDevice } from "../../../../utils/entraDevice.js";
class EntraAdministrativeUnitMemberAddCommand extends GraphCommand {
    get name() {
        return commands.ADMINISTRATIVEUNIT_MEMBER_ADD;
    }
    get description() {
        return 'Adds a member (user, group, device) to an administrative unit';
    }
    constructor() {
        super();
        _EntraAdministrativeUnitMemberAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraAdministrativeUnitMemberAddCommand_instances, "m", _EntraAdministrativeUnitMemberAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraAdministrativeUnitMemberAddCommand_instances, "m", _EntraAdministrativeUnitMemberAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraAdministrativeUnitMemberAddCommand_instances, "m", _EntraAdministrativeUnitMemberAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraAdministrativeUnitMemberAddCommand_instances, "m", _EntraAdministrativeUnitMemberAddCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        let administrativeUnitId = args.options.administrativeUnitId;
        let memberType;
        let memberId;
        try {
            if (args.options.administrativeUnitName) {
                if (this.verbose) {
                    await logger.logToStderr(`Retrieving Administrative Unit Id...`);
                }
                administrativeUnitId = (await entraAdministrativeUnit.getAdministrativeUnitByDisplayName(args.options.administrativeUnitName)).id;
            }
            if (args.options.userId || args.options.userName) {
                memberType = 'users';
                memberId = args.options.userId;
                if (args.options.userName) {
                    if (this.verbose) {
                        await logger.logToStderr(`Retrieving User Id...`);
                    }
                    memberId = await entraUser.getUserIdByUpn(args.options.userName);
                }
            }
            else if (args.options.groupId || args.options.groupName) {
                memberType = 'groups';
                memberId = args.options.groupId;
                if (args.options.groupName) {
                    if (this.verbose) {
                        await logger.logToStderr(`Retrieving Group Id...`);
                    }
                    memberId = await entraGroup.getGroupIdByDisplayName(args.options.groupName);
                }
            }
            else if (args.options.deviceId || args.options.deviceName) {
                memberType = 'devices';
                memberId = args.options.deviceId;
                if (args.options.deviceName) {
                    if (this.verbose) {
                        await logger.logToStderr(`Device with name ${args.options.deviceName} retrieved, returned id: ${memberId}`);
                    }
                    memberId = (await entraDevice.getDeviceByDisplayName(args.options.deviceName)).id;
                }
            }
            const requestOptions = {
                url: `${this.resource}/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/$ref`,
                headers: {
                    'accept': 'application/json;odata.metadata=none'
                },
                data: {
                    "@odata.id": `https://graph.microsoft.com/v1.0/${memberType}/${memberId}`
                }
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_EntraAdministrativeUnitMemberAddCommand_instances = new WeakSet(), _EntraAdministrativeUnitMemberAddCommand_initTelemetry = function _EntraAdministrativeUnitMemberAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            groupId: typeof args.options.groupId !== 'undefined',
            groupName: typeof args.options.groupName !== 'undefined',
            deviceId: typeof args.options.deviceId !== 'undefined',
            deviceName: typeof args.options.deviceName !== 'undefined'
        });
    });
}, _EntraAdministrativeUnitMemberAddCommand_initOptions = function _EntraAdministrativeUnitMemberAddCommand_initOptions() {
    this.options.unshift({
        option: '-i, --administrativeUnitId [administrativeUnitId]'
    }, {
        option: '-n, --administrativeUnitName [administrativeUnitName]'
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
    });
}, _EntraAdministrativeUnitMemberAddCommand_initValidators = function _EntraAdministrativeUnitMemberAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.administrativeUnitId && !validation.isValidGuid(args.options.administrativeUnitId)) {
            return `${args.options.administrativeUnitId} is not a valid GUID`;
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
}, _EntraAdministrativeUnitMemberAddCommand_initOptionSets = function _EntraAdministrativeUnitMemberAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['administrativeUnitId', 'administrativeUnitName'] });
    this.optionSets.push({ options: ['userId', 'userName', 'groupId', 'groupName', 'deviceId', 'deviceName'] });
};
export default new EntraAdministrativeUnitMemberAddCommand();
//# sourceMappingURL=administrativeunit-member-add.js.map