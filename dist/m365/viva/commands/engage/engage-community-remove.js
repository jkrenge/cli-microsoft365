var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _VivaEngageCommunityRemoveCommand_instances, _VivaEngageCommunityRemoveCommand_initTelemetry, _VivaEngageCommunityRemoveCommand_initOptions, _VivaEngageCommunityRemoveCommand_initValidators, _VivaEngageCommunityRemoveCommand_initOptionSets, _VivaEngageCommunityRemoveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import { vivaEngage } from '../../../../utils/vivaEngage.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class VivaEngageCommunityRemoveCommand extends GraphCommand {
    get name() {
        return commands.ENGAGE_COMMUNITY_REMOVE;
    }
    get description() {
        return 'Removes a Viva Engage community';
    }
    constructor() {
        super();
        _VivaEngageCommunityRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _VivaEngageCommunityRemoveCommand_instances, "m", _VivaEngageCommunityRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _VivaEngageCommunityRemoveCommand_instances, "m", _VivaEngageCommunityRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _VivaEngageCommunityRemoveCommand_instances, "m", _VivaEngageCommunityRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _VivaEngageCommunityRemoveCommand_instances, "m", _VivaEngageCommunityRemoveCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _VivaEngageCommunityRemoveCommand_instances, "m", _VivaEngageCommunityRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const removeCommunity = async () => {
            try {
                let communityId = args.options.id;
                if (args.options.displayName) {
                    communityId = (await vivaEngage.getCommunityByDisplayName(args.options.displayName, ['id'])).id;
                }
                else if (args.options.entraGroupId) {
                    communityId = (await vivaEngage.getCommunityByEntraGroupId(args.options.entraGroupId, ['id'])).id;
                }
                if (args.options.verbose) {
                    await logger.logToStderr(`Removing Viva Engage community with ID ${communityId}...`);
                }
                const requestOptions = {
                    url: `${this.resource}/v1.0/employeeExperience/communities/${communityId}`,
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    }
                };
                await request.delete(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeCommunity();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove Viva Engage community '${args.options.id || args.options.displayName || args.options.entraGroupId}'?` });
            if (result) {
                await removeCommunity();
            }
        }
    }
}
_VivaEngageCommunityRemoveCommand_instances = new WeakSet(), _VivaEngageCommunityRemoveCommand_initTelemetry = function _VivaEngageCommunityRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: args.options.id !== 'undefined',
            displayName: args.options.displayName !== 'undefined',
            entraGroupId: args.options.entraGroupId !== 'undefined',
            force: !!args.options.force
        });
    });
}, _VivaEngageCommunityRemoveCommand_initOptions = function _VivaEngageCommunityRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-n, --displayName [displayName]'
    }, {
        option: '--entraGroupId [entraGroupId]'
    }, {
        option: '-f, --force'
    });
}, _VivaEngageCommunityRemoveCommand_initValidators = function _VivaEngageCommunityRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.entraGroupId && !validation.isValidGuid(args.options.entraGroupId)) {
            return `${args.options.entraGroupId} is not a valid GUID for the option 'entraGroupId'.`;
        }
        return true;
    });
}, _VivaEngageCommunityRemoveCommand_initOptionSets = function _VivaEngageCommunityRemoveCommand_initOptionSets() {
    this.optionSets.push({
        options: ['id', 'displayName', 'entraGroupId']
    });
}, _VivaEngageCommunityRemoveCommand_initTypes = function _VivaEngageCommunityRemoveCommand_initTypes() {
    this.types.string.push('id', 'displayName', 'entraGroupId');
};
export default new VivaEngageCommunityRemoveCommand();
//# sourceMappingURL=engage-community-remove.js.map