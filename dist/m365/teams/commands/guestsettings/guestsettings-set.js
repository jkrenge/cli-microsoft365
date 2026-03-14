var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsGuestSettingsSetCommand_instances, _a, _TeamsGuestSettingsSetCommand_initTelemetry, _TeamsGuestSettingsSetCommand_initOptions, _TeamsGuestSettingsSetCommand_initTypes, _TeamsGuestSettingsSetCommand_initValidators;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsGuestSettingsSetCommand extends GraphCommand {
    get name() {
        return commands.GUESTSETTINGS_SET;
    }
    get description() {
        return 'Updates guest settings of a Microsoft Teams team';
    }
    constructor() {
        super();
        _TeamsGuestSettingsSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsGuestSettingsSetCommand_instances, "m", _TeamsGuestSettingsSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsGuestSettingsSetCommand_instances, "m", _TeamsGuestSettingsSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsGuestSettingsSetCommand_instances, "m", _TeamsGuestSettingsSetCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _TeamsGuestSettingsSetCommand_instances, "m", _TeamsGuestSettingsSetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const data = {
            guestSettings: {}
        };
        _a.booleanProps.forEach(p => {
            if (typeof args.options[p] !== 'undefined') {
                data.guestSettings[p] = args.options[p];
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
_a = TeamsGuestSettingsSetCommand, _TeamsGuestSettingsSetCommand_instances = new WeakSet(), _TeamsGuestSettingsSetCommand_initTelemetry = function _TeamsGuestSettingsSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        _a.booleanProps.forEach(p => {
            this.telemetryProperties[p] = args.options[p];
        });
    });
}, _TeamsGuestSettingsSetCommand_initOptions = function _TeamsGuestSettingsSetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --teamId <teamId>'
    }, {
        option: '--allowCreateUpdateChannels [allowCreateUpdateChannels]',
        autocomplete: ['true', 'false']
    }, {
        option: '--allowDeleteChannels [allowDeleteChannels]',
        autocomplete: ['true', 'false']
    });
}, _TeamsGuestSettingsSetCommand_initTypes = function _TeamsGuestSettingsSetCommand_initTypes() {
    this.types.boolean.push('allowCreateUpdateChannels', 'allowDeleteChannels');
}, _TeamsGuestSettingsSetCommand_initValidators = function _TeamsGuestSettingsSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.teamId)) {
            return `${args.options.teamId} is not a valid GUID`;
        }
        return true;
    });
};
TeamsGuestSettingsSetCommand.booleanProps = [
    'allowCreateUpdateChannels',
    'allowDeleteChannels'
];
export default new TeamsGuestSettingsSetCommand();
//# sourceMappingURL=guestsettings-set.js.map