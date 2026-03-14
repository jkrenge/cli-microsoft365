var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsTabRemoveCommand_instances, _TeamsTabRemoveCommand_initTelemetry, _TeamsTabRemoveCommand_initOptions, _TeamsTabRemoveCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsTabRemoveCommand extends GraphCommand {
    get name() {
        return commands.TAB_REMOVE;
    }
    get description() {
        return "Removes a tab from the specified channel";
    }
    constructor() {
        super();
        _TeamsTabRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsTabRemoveCommand_instances, "m", _TeamsTabRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsTabRemoveCommand_instances, "m", _TeamsTabRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsTabRemoveCommand_instances, "m", _TeamsTabRemoveCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const removeTab = async () => {
            const requestOptions = {
                url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(args.options.teamId)}/channels/${args.options.channelId}/tabs/${formatting.encodeQueryParameter(args.options.id)}`,
                headers: {
                    accept: "application/json;odata.metadata=none"
                },
                responseType: 'json'
            };
            try {
                await request.delete(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeTab();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the tab with id ${args.options.id} from channel ${args.options.channelId} in team ${args.options.teamId}?` });
            if (result) {
                await removeTab();
            }
        }
    }
}
_TeamsTabRemoveCommand_instances = new WeakSet(), _TeamsTabRemoveCommand_initTelemetry = function _TeamsTabRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: (!!args.options.force).toString()
        });
    });
}, _TeamsTabRemoveCommand_initOptions = function _TeamsTabRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-t, --teamId <teamId>'
    }, {
        option: '-c, --channelId <channelId>'
    }, {
        option: '-i, --id <id>'
    }, {
        option: '-f, --force'
    });
}, _TeamsTabRemoveCommand_initValidators = function _TeamsTabRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.teamId)) {
            return `${args.options.teamId} is not a valid GUID`;
        }
        if (!validation.isValidTeamsChannelId(args.options.channelId)) {
            return `${args.options.channelId} is not a valid Teams ChannelId`;
        }
        if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
};
export default new TeamsTabRemoveCommand();
//# sourceMappingURL=tab-remove.js.map