var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsMessagingSettingsSetCommand_instances, _a, _TeamsMessagingSettingsSetCommand_initTelemetry, _TeamsMessagingSettingsSetCommand_initOptions, _TeamsMessagingSettingsSetCommand_initTypes, _TeamsMessagingSettingsSetCommand_initValidators;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsMessagingSettingsSetCommand extends GraphCommand {
    get name() {
        return commands.MESSAGINGSETTINGS_SET;
    }
    get description() {
        return 'Updates messaging settings of a Microsoft Teams team';
    }
    constructor() {
        super();
        _TeamsMessagingSettingsSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsMessagingSettingsSetCommand_instances, "m", _TeamsMessagingSettingsSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsMessagingSettingsSetCommand_instances, "m", _TeamsMessagingSettingsSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsMessagingSettingsSetCommand_instances, "m", _TeamsMessagingSettingsSetCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _TeamsMessagingSettingsSetCommand_instances, "m", _TeamsMessagingSettingsSetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const data = {
            messagingSettings: {}
        };
        _a.booleanProps.forEach((p) => {
            if (typeof args.options[p] !== 'undefined') {
                data.messagingSettings[p] = args.options[p];
            }
        });
        const requestOptions = {
            url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(args.options.teamId)}`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            data: data,
            responseType: 'json'
        };
        try {
            await request.patch(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_a = TeamsMessagingSettingsSetCommand, _TeamsMessagingSettingsSetCommand_instances = new WeakSet(), _TeamsMessagingSettingsSetCommand_initTelemetry = function _TeamsMessagingSettingsSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        _a.booleanProps.forEach((p) => {
            this.telemetryProperties[p] = args.options[p];
        });
    });
}, _TeamsMessagingSettingsSetCommand_initOptions = function _TeamsMessagingSettingsSetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --teamId <teamId>'
    }, {
        option: '--allowUserEditMessages [allowUserEditMessages]',
        autocomplete: ['true', 'false']
    }, {
        option: '--allowUserDeleteMessages [allowUserDeleteMessages]',
        autocomplete: ['true', 'false']
    }, {
        option: '--allowOwnerDeleteMessages [allowOwnerDeleteMessages]',
        autocomplete: ['true', 'false']
    }, {
        option: '--allowTeamMentions [allowTeamMentions]',
        autocomplete: ['true', 'false']
    }, {
        option: '--allowChannelMentions [allowChannelMentions]',
        autocomplete: ['true', 'false']
    });
}, _TeamsMessagingSettingsSetCommand_initTypes = function _TeamsMessagingSettingsSetCommand_initTypes() {
    this.types.boolean.push('allowUserEditMessages', 'allowUserDeleteMessages', 'allowOwnerDeleteMessages', 'allowTeamMentions', 'allowChannelMentions');
}, _TeamsMessagingSettingsSetCommand_initValidators = function _TeamsMessagingSettingsSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.teamId)) {
            return `${args.options.teamId} is not a valid GUID`;
        }
        let hasDuplicate = false;
        let property = '';
        _a.booleanProps.forEach((prop) => {
            if (args.options[prop] instanceof Array) {
                property = prop;
                hasDuplicate = true;
            }
        });
        if (hasDuplicate) {
            return `Duplicate option ${property} specified. Specify only one`;
        }
        return true;
    });
};
TeamsMessagingSettingsSetCommand.booleanProps = [
    'allowUserEditMessages',
    'allowUserDeleteMessages',
    'allowOwnerDeleteMessages',
    'allowTeamMentions',
    'allowChannelMentions'
];
export default new TeamsMessagingSettingsSetCommand();
//# sourceMappingURL=messagingsettings-set.js.map