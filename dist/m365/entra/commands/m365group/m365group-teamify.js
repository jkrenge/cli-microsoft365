var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraM365GroupTeamifyCommand_instances, _EntraM365GroupTeamifyCommand_initTelemetry, _EntraM365GroupTeamifyCommand_initOptions, _EntraM365GroupTeamifyCommand_initValidators, _EntraM365GroupTeamifyCommand_initOptionSets, _EntraM365GroupTeamifyCommand_initTypes;
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraM365GroupTeamifyCommand extends GraphCommand {
    get name() {
        return commands.M365GROUP_TEAMIFY;
    }
    get description() {
        return 'Creates a new Microsoft Teams team under existing Microsoft 365 group';
    }
    constructor() {
        super();
        _EntraM365GroupTeamifyCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraM365GroupTeamifyCommand_instances, "m", _EntraM365GroupTeamifyCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupTeamifyCommand_instances, "m", _EntraM365GroupTeamifyCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupTeamifyCommand_instances, "m", _EntraM365GroupTeamifyCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupTeamifyCommand_instances, "m", _EntraM365GroupTeamifyCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupTeamifyCommand_instances, "m", _EntraM365GroupTeamifyCommand_initTypes).call(this);
    }
    async getGroupId(options) {
        if (options.id) {
            return options.id;
        }
        if (options.displayName) {
            return await entraGroup.getGroupIdByDisplayName(options.displayName);
        }
        return await entraGroup.getGroupIdByMailNickname(options.mailNickname);
    }
    async commandAction(logger, args) {
        try {
            const groupId = await this.getGroupId(args.options);
            const isUnifiedGroup = await entraGroup.isUnifiedGroup(groupId);
            if (!isUnifiedGroup) {
                throw Error(`Specified group with id '${groupId}' is not a Microsoft 365 group.`);
            }
            const data = {
                "memberSettings": {
                    "allowCreatePrivateChannels": true,
                    "allowCreateUpdateChannels": true
                },
                "messagingSettings": {
                    "allowUserEditMessages": true,
                    "allowUserDeleteMessages": true
                },
                "funSettings": {
                    "allowGiphy": true,
                    "giphyContentRating": "strict"
                }
            };
            const requestOptions = {
                url: `${this.resource}/v1.0/groups/${formatting.encodeQueryParameter(groupId)}/team`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                data: data,
                responseType: 'json'
            };
            await request.put(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_EntraM365GroupTeamifyCommand_instances = new WeakSet(), _EntraM365GroupTeamifyCommand_initTelemetry = function _EntraM365GroupTeamifyCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            mailNickname: typeof args.options.mailNickname !== 'undefined',
            displayName: typeof args.options.displayName !== 'undefined'
        });
    });
}, _EntraM365GroupTeamifyCommand_initOptions = function _EntraM365GroupTeamifyCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-n, --displayName [displayName]'
    }, {
        option: '--mailNickname [mailNickname]'
    });
}, _EntraM365GroupTeamifyCommand_initValidators = function _EntraM365GroupTeamifyCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
}, _EntraM365GroupTeamifyCommand_initOptionSets = function _EntraM365GroupTeamifyCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'displayName', 'mailNickname'] });
}, _EntraM365GroupTeamifyCommand_initTypes = function _EntraM365GroupTeamifyCommand_initTypes() {
    this.types.string.push('id', 'displayName', 'mailNickname');
};
export default new EntraM365GroupTeamifyCommand();
//# sourceMappingURL=m365group-teamify.js.map