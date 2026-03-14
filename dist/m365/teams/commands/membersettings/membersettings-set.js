var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsMemberSettingsSetCommand_instances, _a, _TeamsMemberSettingsSetCommand_initTelemetry, _TeamsMemberSettingsSetCommand_initOptions, _TeamsMemberSettingsSetCommand_initTypes, _TeamsMemberSettingsSetCommand_initValidators;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsMemberSettingsSetCommand extends GraphCommand {
    get name() {
        return commands.MEMBERSETTINGS_SET;
    }
    get description() {
        return 'Updates member settings of a Microsoft Teams team';
    }
    constructor() {
        super();
        _TeamsMemberSettingsSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsMemberSettingsSetCommand_instances, "m", _TeamsMemberSettingsSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsMemberSettingsSetCommand_instances, "m", _TeamsMemberSettingsSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsMemberSettingsSetCommand_instances, "m", _TeamsMemberSettingsSetCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _TeamsMemberSettingsSetCommand_instances, "m", _TeamsMemberSettingsSetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const data = {
            memberSettings: {}
        };
        _a.booleanProps.forEach(p => {
            if (typeof args.options[p] !== 'undefined') {
                data.memberSettings[p] = args.options[p];
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
_a = TeamsMemberSettingsSetCommand, _TeamsMemberSettingsSetCommand_instances = new WeakSet(), _TeamsMemberSettingsSetCommand_initTelemetry = function _TeamsMemberSettingsSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        _a.booleanProps.forEach(p => {
            this.telemetryProperties[p] = args.options[p];
        });
    });
}, _TeamsMemberSettingsSetCommand_initOptions = function _TeamsMemberSettingsSetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --teamId <teamId>'
    }, {
        option: '--allowAddRemoveApps [allowAddRemoveApps]',
        autocomplete: ['true', 'false']
    }, {
        option: '--allowCreateUpdateChannels [allowCreateUpdateChannels]',
        autocomplete: ['true', 'false']
    }, {
        option: '--allowCreateUpdateRemoveConnectors [allowCreateUpdateRemoveConnectors]',
        autocomplete: ['true', 'false']
    }, {
        option: '--allowCreateUpdateRemoveTabs [allowCreateUpdateRemoveTabs]',
        autocomplete: ['true', 'false']
    }, {
        option: '--allowDeleteChannels [allowDeleteChannels]',
        autocomplete: ['true', 'false']
    });
}, _TeamsMemberSettingsSetCommand_initTypes = function _TeamsMemberSettingsSetCommand_initTypes() {
    this.types.boolean.push('allowAddRemoveApps', 'allowCreateUpdateChannels', 'allowCreateUpdateRemoveConnectors', 'allowCreateUpdateRemoveTabs', 'allowDeleteChannels');
}, _TeamsMemberSettingsSetCommand_initValidators = function _TeamsMemberSettingsSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.teamId)) {
            return `${args.options.teamId} is not a valid GUID`;
        }
        return true;
    });
};
TeamsMemberSettingsSetCommand.booleanProps = [
    'allowAddRemoveApps',
    'allowCreateUpdateChannels',
    'allowCreateUpdateRemoveConnectors',
    'allowCreateUpdateRemoveTabs',
    'allowDeleteChannels'
];
export default new TeamsMemberSettingsSetCommand();
//# sourceMappingURL=membersettings-set.js.map