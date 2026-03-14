var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _VivaEngageCommunityAddCommand_instances, _VivaEngageCommunityAddCommand_initTelemetry, _VivaEngageCommunityAddCommand_initOptions, _VivaEngageCommunityAddCommand_initValidators, _VivaEngageCommunityAddCommand_initTypes, _VivaEngageCommunityAddCommand_initOptionSets;
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import { accessToken } from '../../../../utils/accessToken.js';
import auth from '../../../../Auth.js';
import { formatting } from '../../../../utils/formatting.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { setTimeout } from 'timers/promises';
class VivaEngageCommunityAddCommand extends GraphCommand {
    get name() {
        return commands.ENGAGE_COMMUNITY_ADD;
    }
    get description() {
        return 'Creates a new community in Viva Engage';
    }
    constructor() {
        super();
        _VivaEngageCommunityAddCommand_instances.add(this);
        this.pollingInterval = 5000;
        this.privacyOptions = ['public', 'private'];
        __classPrivateFieldGet(this, _VivaEngageCommunityAddCommand_instances, "m", _VivaEngageCommunityAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _VivaEngageCommunityAddCommand_instances, "m", _VivaEngageCommunityAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _VivaEngageCommunityAddCommand_instances, "m", _VivaEngageCommunityAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _VivaEngageCommunityAddCommand_instances, "m", _VivaEngageCommunityAddCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _VivaEngageCommunityAddCommand_instances, "m", _VivaEngageCommunityAddCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const { displayName, description, privacy, adminEntraIds, adminEntraUserNames, wait } = args.options;
        const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken);
        if (isAppOnlyAccessToken && !adminEntraIds && !adminEntraUserNames) {
            this.handleError(`Specify at least one admin using either adminEntraIds or adminEntraUserNames options when using application permissions.`);
        }
        if (this.verbose) {
            await logger.logToStderr(`Creating a Viva Engage community with display name '${displayName}'...`);
        }
        try {
            const requestOptions = {
                url: `${this.resource}/beta/employeeExperience/communities`,
                headers: {
                    accept: 'application/json;odata.metadata=none',
                    'content-type': 'application/json'
                },
                responseType: 'json',
                fullResponse: true,
                data: {
                    displayName: displayName,
                    description: description,
                    privacy: privacy
                }
            };
            const entraIds = await this.getGraphUserUrls(args.options);
            if (entraIds.length > 0) {
                requestOptions.data['owners@odata.bind'] = entraIds;
            }
            const res = await request.post(requestOptions);
            const location = res.headers.location;
            if (!wait) {
                await logger.log(location);
                return;
            }
            let status;
            do {
                if (this.verbose) {
                    await logger.logToStderr(`Community still provisioning. Retrying in ${this.pollingInterval / 1000} seconds...`);
                }
                await setTimeout(this.pollingInterval);
                if (this.verbose) {
                    await logger.logToStderr(`Checking create community operation status...`);
                }
                const operation = await request.get({
                    url: location,
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    },
                    responseType: 'json'
                });
                status = operation.status;
                if (this.verbose) {
                    await logger.logToStderr(`Community creation operation status: ${status}`);
                }
                if (status === 'failed') {
                    throw `Community creation failed: ${operation.statusDetail}`;
                }
                if (status === 'succeeded') {
                    await logger.log(operation);
                }
            } while (status === 'notStarted' || status === 'running');
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getGraphUserUrls(options) {
        let entraIds = [];
        if (options.adminEntraIds) {
            entraIds = formatting.splitAndTrim(options.adminEntraIds);
        }
        else if (options.adminEntraUserNames) {
            entraIds = await entraUser.getUserIdsByUpns(formatting.splitAndTrim(options.adminEntraUserNames));
        }
        const graphUserUrls = entraIds.map(id => `${this.resource}/beta/users/${id}`);
        return graphUserUrls;
    }
}
_VivaEngageCommunityAddCommand_instances = new WeakSet(), _VivaEngageCommunityAddCommand_initTelemetry = function _VivaEngageCommunityAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            adminEntraIds: typeof args.options.adminEntraIds !== 'undefined',
            adminEntraUserNames: typeof args.options.adminEntraUserNames !== 'undefined',
            wait: !!args.options.wait
        });
    });
}, _VivaEngageCommunityAddCommand_initOptions = function _VivaEngageCommunityAddCommand_initOptions() {
    this.options.unshift({ option: '--displayName <displayName>' }, { option: '--description <description>' }, {
        option: '--privacy <privacy>',
        autocomplete: this.privacyOptions
    }, { option: '--adminEntraIds [adminEntraIds]' }, { option: '--adminEntraUserNames [adminEntraUserNames]' }, { option: '--wait' });
}, _VivaEngageCommunityAddCommand_initValidators = function _VivaEngageCommunityAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.displayName.length > 255) {
            return `The maximum amount of characters for 'displayName' is 255.`;
        }
        if (args.options.description.length > 1024) {
            return `The maximum amount of characters for 'description' is 1024.`;
        }
        if (this.privacyOptions.indexOf(args.options.privacy) === -1) {
            return `'${args.options.privacy}' is not a valid value for privacy. Allowed values are: ${this.privacyOptions.join(', ')}.`;
        }
        if (args.options.adminEntraIds) {
            const isValidGUIDArrayResult = validation.isValidGuidArray(args.options.adminEntraIds);
            if (isValidGUIDArrayResult !== true) {
                return `The following GUIDs are invalid for the option 'adminEntraIds': ${isValidGUIDArrayResult}.`;
            }
            if (formatting.splitAndTrim(args.options.adminEntraIds).length > 20) {
                return `Maximum of 20 admins allowed. Please reduce the number of users and try again.`;
            }
        }
        if (args.options.adminEntraUserNames) {
            const isValidUPNArrayResult = validation.isValidUserPrincipalNameArray(args.options.adminEntraUserNames);
            if (isValidUPNArrayResult !== true) {
                return `The following user principal names are invalid for the option 'adminEntraUserNames': ${isValidUPNArrayResult}.`;
            }
            if (formatting.splitAndTrim(args.options.adminEntraUserNames).length > 20) {
                return `Maximum of 20 admins allowed. Please reduce the number of users and try again.`;
            }
        }
        return true;
    });
}, _VivaEngageCommunityAddCommand_initTypes = function _VivaEngageCommunityAddCommand_initTypes() {
    this.types.string.push('displayName', 'description', 'privacy', 'adminEntraIds', 'adminEntraUserNames');
    this.types.boolean.push('wait');
}, _VivaEngageCommunityAddCommand_initOptionSets = function _VivaEngageCommunityAddCommand_initOptionSets() {
    this.optionSets.push({
        options: ['adminEntraIds', 'adminEntraUserNames'],
        runsWhen: (args) => args.options.adminEntraIds || args.options.adminEntraUserNames
    });
};
export default new VivaEngageCommunityAddCommand();
//# sourceMappingURL=engage-community-add.js.map