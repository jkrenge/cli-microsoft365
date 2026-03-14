var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsTeamSetCommand_instances, _a, _TeamsTeamSetCommand_initTelemetry, _TeamsTeamSetCommand_initOptions, _TeamsTeamSetCommand_initValidators;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsTeamSetCommand extends GraphCommand {
    get name() {
        return commands.TEAM_SET;
    }
    get description() {
        return 'Updates settings of a Microsoft Teams team';
    }
    constructor() {
        super();
        _TeamsTeamSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsTeamSetCommand_instances, "m", _TeamsTeamSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsTeamSetCommand_instances, "m", _TeamsTeamSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsTeamSetCommand_instances, "m", _TeamsTeamSetCommand_initValidators).call(this);
    }
    mapRequestBody(options) {
        const requestBody = {};
        if (options.name) {
            requestBody.displayName = options.name;
        }
        if (options.description) {
            requestBody.description = options.description;
        }
        if (options.mailNickName) {
            requestBody.mailNickName = options.mailNickName;
        }
        if (options.classification) {
            requestBody.classification = options.classification;
        }
        if (options.visibility) {
            requestBody.visibility = options.visibility;
        }
        return requestBody;
    }
    async commandAction(logger, args) {
        const data = this.mapRequestBody(args.options);
        const requestOptions = {
            url: `${this.resource}/v1.0/groups/${formatting.encodeQueryParameter(args.options.id)}`,
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
_a = TeamsTeamSetCommand, _TeamsTeamSetCommand_instances = new WeakSet(), _TeamsTeamSetCommand_initTelemetry = function _TeamsTeamSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        _a.props.forEach((p) => {
            this.telemetryProperties[p] = typeof args.options[p] !== 'undefined';
        });
    });
}, _TeamsTeamSetCommand_initOptions = function _TeamsTeamSetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-n, --name [name]'
    }, {
        option: '--description [description]'
    }, {
        option: '--mailNickName [mailNickName]'
    }, {
        option: '--classification [classification]'
    }, {
        option: '--visibility [visibility]',
        autocomplete: ['Private', 'Public']
    });
}, _TeamsTeamSetCommand_initValidators = function _TeamsTeamSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        if (args.options.visibility) {
            if (args.options.visibility.toLowerCase() !== 'private' && args.options.visibility.toLowerCase() !== 'public') {
                return `${args.options.visibility} is not a valid visibility type. Allowed values are Private|Public`;
            }
        }
        return true;
    });
};
TeamsTeamSetCommand.props = [
    'description',
    'mailNickName',
    'classification',
    'visibility '
];
export default new TeamsTeamSetCommand();
//# sourceMappingURL=team-set.js.map