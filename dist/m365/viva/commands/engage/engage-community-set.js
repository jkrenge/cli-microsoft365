var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _VivaEngageCommunitySetCommand_instances, _VivaEngageCommunitySetCommand_initTelemetry, _VivaEngageCommunitySetCommand_initOptions, _VivaEngageCommunitySetCommand_initValidators, _VivaEngageCommunitySetCommand_initTypes, _VivaEngageCommunitySetCommand_initOptionSets;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import { vivaEngage } from '../../../../utils/vivaEngage.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class VivaEngageCommunitySetCommand extends GraphCommand {
    get name() {
        return commands.ENGAGE_COMMUNITY_SET;
    }
    get description() {
        return 'Updates an existing Viva Engage community';
    }
    constructor() {
        super();
        _VivaEngageCommunitySetCommand_instances.add(this);
        this.privacyOptions = ['public', 'private'];
        __classPrivateFieldGet(this, _VivaEngageCommunitySetCommand_instances, "m", _VivaEngageCommunitySetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _VivaEngageCommunitySetCommand_instances, "m", _VivaEngageCommunitySetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _VivaEngageCommunitySetCommand_instances, "m", _VivaEngageCommunitySetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _VivaEngageCommunitySetCommand_instances, "m", _VivaEngageCommunitySetCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _VivaEngageCommunitySetCommand_instances, "m", _VivaEngageCommunitySetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        let communityId = args.options.id;
        if (args.options.displayName) {
            communityId = (await vivaEngage.getCommunityByDisplayName(args.options.displayName, ['id'])).id;
        }
        else if (args.options.entraGroupId) {
            communityId = (await vivaEngage.getCommunityByEntraGroupId(args.options.entraGroupId, ['id'])).id;
        }
        if (this.verbose) {
            await logger.logToStderr(`Updating Viva Engage community with ID ${communityId}...`);
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/employeeExperience/communities/${communityId}`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json',
            data: {
                description: args.options.description,
                displayName: args.options.newDisplayName,
                privacy: args.options.privacy
            }
        };
        try {
            await request.patch(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_VivaEngageCommunitySetCommand_instances = new WeakSet(), _VivaEngageCommunitySetCommand_initTelemetry = function _VivaEngageCommunitySetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            displayName: typeof args.options.displayName !== 'undefined',
            entraGroupId: typeof args.options.entraGroupId !== 'undefined',
            newDisplayName: typeof args.options.newDisplayName !== 'undefined',
            description: typeof args.options.description !== 'undefined',
            privacy: typeof args.options.privacy !== 'undefined'
        });
    });
}, _VivaEngageCommunitySetCommand_initOptions = function _VivaEngageCommunitySetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-d, --displayName [displayName]'
    }, {
        option: '--entraGroupId [entraGroupId]'
    }, {
        option: '--newDisplayName [newDisplayName]'
    }, {
        option: '--description [description]'
    }, {
        option: '--privacy [privacy]',
        autocomplete: this.privacyOptions
    });
}, _VivaEngageCommunitySetCommand_initValidators = function _VivaEngageCommunitySetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.entraGroupId && !validation.isValidGuid(args.options.entraGroupId)) {
            return `${args.options.entraGroupId} is not a valid GUID for the option 'entraGroupId'.`;
        }
        if (args.options.newDisplayName && args.options.newDisplayName.length > 255) {
            return `The maximum amount of characters for 'newDisplayName' is 255.`;
        }
        if (args.options.description && args.options.description.length > 1024) {
            return `The maximum amount of characters for 'description' is 1024.`;
        }
        if (args.options.privacy && this.privacyOptions.map(x => x.toLowerCase()).indexOf(args.options.privacy.toLowerCase()) === -1) {
            return `${args.options.privacy} is not a valid privacy. Allowed values are ${this.privacyOptions.join(', ')}`;
        }
        if (!args.options.newDisplayName && !args.options.description && !args.options.privacy) {
            return 'Specify at least newDisplayName, description, or privacy.';
        }
        return true;
    });
}, _VivaEngageCommunitySetCommand_initTypes = function _VivaEngageCommunitySetCommand_initTypes() {
    this.types.string.push('id', 'displayName', 'entraGroupId', 'newDisplayName', 'description', 'privacy');
}, _VivaEngageCommunitySetCommand_initOptionSets = function _VivaEngageCommunitySetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'displayName', 'entraGroupId'] });
};
export default new VivaEngageCommunitySetCommand();
//# sourceMappingURL=engage-community-set.js.map