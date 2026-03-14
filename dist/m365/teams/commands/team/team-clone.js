var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsTeamCloneCommand_instances, _TeamsTeamCloneCommand_initTelemetry, _TeamsTeamCloneCommand_initOptions, _TeamsTeamCloneCommand_initValidators;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsTeamCloneCommand extends GraphCommand {
    get name() {
        return commands.TEAM_CLONE;
    }
    get description() {
        return 'Creates a clone of a Microsoft Teams team';
    }
    constructor() {
        super();
        _TeamsTeamCloneCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsTeamCloneCommand_instances, "m", _TeamsTeamCloneCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsTeamCloneCommand_instances, "m", _TeamsTeamCloneCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsTeamCloneCommand_instances, "m", _TeamsTeamCloneCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const data = {
            displayName: args.options.name,
            mailNickname: this.generateMailNickname(args.options.name),
            partsToClone: args.options.partsToClone
        };
        if (args.options.description) {
            data.description = args.options.description;
        }
        if (args.options.classification) {
            data.classification = args.options.classification;
        }
        if (args.options.visibility) {
            data.visibility = args.options.visibility;
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(args.options.id)}/clone`,
            headers: {
                'content-type': 'application/json',
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json',
            data: data
        };
        try {
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    /**
     * There is a know issue that the mailNickname is currently ignored and cannot be set by the user
     * However the mailNickname is still required by the payload so to deliver better user experience
     * the CLI generates mailNickname for the user
     * so the user does not have to specify something that will be ignored.
     * For more see: https://docs.microsoft.com/en-us/graph/api/team-clone?view=graph-rest-1.0#request-data
     * This method has to be removed once the graph team fixes the issue and then the actual value
     * of the mailNickname would have to be specified by the CLI user.
     * @param displayName teams display name
     */
    generateMailNickname(displayName) {
        // currently the Microsoft Graph generates mailNickname in a similar fashion
        return `${displayName.replace(/[^a-zA-Z0-9]/g, "")}${Math.floor(Math.random() * 9999)}`;
    }
}
_TeamsTeamCloneCommand_instances = new WeakSet(), _TeamsTeamCloneCommand_initTelemetry = function _TeamsTeamCloneCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            description: typeof args.options.description !== 'undefined',
            classification: typeof args.options.classification !== 'undefined',
            visibility: typeof args.options.visibility !== 'undefined'
        });
    });
}, _TeamsTeamCloneCommand_initOptions = function _TeamsTeamCloneCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '-n, --name <name>'
    }, {
        option: '-p, --partsToClone <partsToClone>',
        autocomplete: ['apps', 'channels', 'members', 'settings', 'tabs']
    }, {
        option: '-d, --description [description]'
    }, {
        option: '-c, --classification [classification]'
    }, {
        option: '-v, --visibility [visibility]',
        autocomplete: ['Private', 'Public']
    });
}, _TeamsTeamCloneCommand_initValidators = function _TeamsTeamCloneCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        const partsToClone = args.options.partsToClone.replace(/\s/g, '').split(',');
        for (const partToClone of partsToClone) {
            const part = partToClone.toLowerCase();
            if (part !== 'apps' &&
                part !== 'channels' &&
                part !== 'members' &&
                part !== 'settings' &&
                part !== 'tabs') {
                return `${part} is not a valid partsToClone. Allowed values are apps|channels|members|settings|tabs`;
            }
        }
        if (args.options.visibility) {
            const visibility = args.options.visibility.toLowerCase();
            if (visibility !== 'private' &&
                visibility !== 'public') {
                return `${args.options.visibility} is not a valid visibility type. Allowed values are Private|Public`;
            }
        }
        return true;
    });
};
export default new TeamsTeamCloneCommand();
//# sourceMappingURL=team-clone.js.map