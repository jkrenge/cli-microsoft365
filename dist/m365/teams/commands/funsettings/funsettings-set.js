var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsFunSettingsSetCommand_instances, _a, _TeamsFunSettingsSetCommand_initTelemetry, _TeamsFunSettingsSetCommand_initOptions, _TeamsFunSettingsSetCommand_initTypes, _TeamsFunSettingsSetCommand_initValidators;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsFunSettingsSetCommand extends GraphCommand {
    get name() {
        return commands.FUNSETTINGS_SET;
    }
    get description() {
        return 'Updates fun settings of a Microsoft Teams team';
    }
    constructor() {
        super();
        _TeamsFunSettingsSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsFunSettingsSetCommand_instances, "m", _TeamsFunSettingsSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsFunSettingsSetCommand_instances, "m", _TeamsFunSettingsSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsFunSettingsSetCommand_instances, "m", _TeamsFunSettingsSetCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _TeamsFunSettingsSetCommand_instances, "m", _TeamsFunSettingsSetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Updating fun settings for team ${args.options.teamId}`);
            }
            const data = {
                funSettings: {}
            };
            _a.booleanProps.forEach(p => {
                if (typeof args.options[p] !== 'undefined') {
                    data.funSettings[p] = args.options[p];
                }
            });
            if (args.options.giphyContentRating) {
                data.funSettings.giphyContentRating = args.options.giphyContentRating;
            }
            const requestOptions = {
                url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(args.options.teamId)}`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                data: data,
                responseType: 'json'
            };
            await request.patch(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_a = TeamsFunSettingsSetCommand, _TeamsFunSettingsSetCommand_instances = new WeakSet(), _TeamsFunSettingsSetCommand_initTelemetry = function _TeamsFunSettingsSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            giphyContentRating: args.options.giphyContentRating
        });
        _a.booleanProps.forEach(p => {
            this.telemetryProperties[p] = args.options[p];
        });
    });
}, _TeamsFunSettingsSetCommand_initOptions = function _TeamsFunSettingsSetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --teamId <teamId>'
    }, {
        option: '--allowGiphy [allowGiphy]',
        autocomplete: ['true', 'false']
    }, {
        option: '--giphyContentRating [giphyContentRating]'
    }, {
        option: '--allowStickersAndMemes [allowStickersAndMemes]',
        autocomplete: ['true', 'false']
    }, {
        option: '--allowCustomMemes [allowCustomMemes]',
        autocomplete: ['true', 'false']
    });
}, _TeamsFunSettingsSetCommand_initTypes = function _TeamsFunSettingsSetCommand_initTypes() {
    this.types.boolean.push('allowGiphy', 'allowStickersAndMemes', 'allowCustomMemes');
}, _TeamsFunSettingsSetCommand_initValidators = function _TeamsFunSettingsSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.teamId)) {
            return `${args.options.teamId} is not a valid GUID`;
        }
        if (args.options.giphyContentRating) {
            const giphyContentRating = args.options.giphyContentRating.toLowerCase();
            if (giphyContentRating !== 'strict' && giphyContentRating !== 'moderate') {
                return `giphyContentRating value ${args.options.giphyContentRating} is not valid.  Please specify Strict or Moderate.`;
            }
        }
        return true;
    });
};
TeamsFunSettingsSetCommand.booleanProps = [
    'allowGiphy',
    'allowStickersAndMemes',
    'allowCustomMemes'
];
export default new TeamsFunSettingsSetCommand();
//# sourceMappingURL=funsettings-set.js.map