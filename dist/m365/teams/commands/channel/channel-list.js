var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsChannelListCommand_instances, _TeamsChannelListCommand_initTelemetry, _TeamsChannelListCommand_initOptions, _TeamsChannelListCommand_initValidators, _TeamsChannelListCommand_initOptionSets;
import { entraGroup } from '../../../../utils/entraGroup.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsChannelListCommand extends GraphCommand {
    get name() {
        return commands.CHANNEL_LIST;
    }
    get description() {
        return 'Lists channels in the specified Microsoft Teams team';
    }
    defaultProperties() {
        return ['id', 'displayName'];
    }
    constructor() {
        super();
        _TeamsChannelListCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsChannelListCommand_instances, "m", _TeamsChannelListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsChannelListCommand_instances, "m", _TeamsChannelListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsChannelListCommand_instances, "m", _TeamsChannelListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsChannelListCommand_instances, "m", _TeamsChannelListCommand_initOptionSets).call(this);
    }
    async getTeamId(args) {
        if (args.options.teamId) {
            return args.options.teamId;
        }
        const group = await entraGroup.getGroupByDisplayName(args.options.teamName);
        if (group.resourceProvisioningOptions.indexOf('Team') === -1) {
            throw 'The specified team does not exist in the Microsoft Teams';
        }
        return group.id;
    }
    async commandAction(logger, args) {
        try {
            const teamId = await this.getTeamId(args);
            let endpoint = `${this.resource}/v1.0/teams/${teamId}/channels`;
            if (args.options.type) {
                endpoint += `?$filter=membershipType eq '${args.options.type}'`;
            }
            const items = await odata.getAllItems(endpoint);
            await logger.log(items);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_TeamsChannelListCommand_instances = new WeakSet(), _TeamsChannelListCommand_initTelemetry = function _TeamsChannelListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            teamId: typeof args.options.teamId !== 'undefined',
            teamName: typeof args.options.teamName !== 'undefined'
        });
    });
}, _TeamsChannelListCommand_initOptions = function _TeamsChannelListCommand_initOptions() {
    this.options.unshift({
        option: '-i, --teamId [teamId]'
    }, {
        option: '--teamName [teamName]'
    }, {
        option: '--type [type]',
        autocomplete: ['standard', 'private', 'shared']
    });
}, _TeamsChannelListCommand_initValidators = function _TeamsChannelListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
            return `${args.options.teamId} is not a valid GUID`;
        }
        if (args.options.type && ['standard', 'private', 'shared'].indexOf(args.options.type.toLowerCase()) === -1) {
            return `${args.options.type} is not a valid type value. Allowed values standard|private|shared`;
        }
        return true;
    });
}, _TeamsChannelListCommand_initOptionSets = function _TeamsChannelListCommand_initOptionSets() {
    this.optionSets.push({ options: ['teamId', 'teamName'] });
};
export default new TeamsChannelListCommand();
//# sourceMappingURL=channel-list.js.map