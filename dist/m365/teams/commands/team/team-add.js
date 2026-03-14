var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsTeamAddCommand_instances, _TeamsTeamAddCommand_initTelemetry, _TeamsTeamAddCommand_initOptions, _TeamsTeamAddCommand_initValidators, _TeamsTeamAddCommand_initOptionSets;
import { setTimeout } from 'timers/promises';
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { accessToken } from '../../../../utils/accessToken.js';
import auth from '../../../../Auth.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { formatting } from '../../../../utils/formatting.js';
class TeamsTeamAddCommand extends GraphCommand {
    get name() {
        return commands.TEAM_ADD;
    }
    get description() {
        return 'Adds a new Microsoft Teams team';
    }
    constructor() {
        super();
        _TeamsTeamAddCommand_instances.add(this);
        this.pollingInterval = 30000;
        __classPrivateFieldGet(this, _TeamsTeamAddCommand_instances, "m", _TeamsTeamAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsTeamAddCommand_instances, "m", _TeamsTeamAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsTeamAddCommand_instances, "m", _TeamsTeamAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsTeamAddCommand_instances, "m", _TeamsTeamAddCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken);
        if (isAppOnlyAccessToken && !args.options.ownerUserNames && !args.options.ownerIds && !args.options.ownerEmails) {
            this.handleError(`Specify at least 'ownerUserNames', 'ownerIds' or 'ownerEmails' when using application permissions.`);
        }
        let requestBody;
        if (args.options.template) {
            if (this.verbose) {
                await logger.logToStderr(`Using template...`);
            }
            requestBody = JSON.parse(args.options.template);
            if (args.options.name) {
                if (this.verbose) {
                    await logger.logToStderr(`Using '${args.options.name}' as name...`);
                }
                requestBody.displayName = args.options.name;
            }
            if (args.options.description) {
                if (this.verbose) {
                    await logger.logToStderr(`Using '${args.options.description}' as description...`);
                }
                requestBody.description = args.options.description;
            }
        }
        else {
            if (this.verbose) {
                await logger.logToStderr(`Creating team with name ${args.options.name} and description ${args.options.description}`);
            }
            requestBody = {
                'template@odata.bind': `https://graph.microsoft.com/v1.0/teamsTemplates('standard')`,
                displayName: args.options.name,
                description: args.options.description
            };
        }
        let members = [];
        if (args.options.ownerEmails || args.options.ownerIds || args.options.ownerUserNames) {
            await this.retrieveMembersToAdd(members, 'owner', args.options.ownerEmails, args.options.ownerIds, args.options.ownerUserNames);
        }
        if (args.options.memberEmails || args.options.memberIds || args.options.memberUserNames) {
            await this.retrieveMembersToAdd(members, 'member', args.options.memberEmails, args.options.memberIds, args.options.memberUserNames);
        }
        // We filter out the first owner here and add it to the request body when we are using application only permissions. This is required or the Graph API will throw an error.
        if (members.length > 0 && members.filter(y => y.roles.includes('owner')).length > 0 && isAppOnlyAccessToken) {
            const groupOwner = members.filter(y => y.roles.includes('owner')).slice(0, 1);
            members = members.filter(y => y !== groupOwner[0]);
            requestBody.members = groupOwner;
        }
        const requestOptionsPost = {
            url: `${this.resource}/v1.0/teams`,
            headers: {
                'accept': 'application/json;odata.metadata=none'
            },
            data: requestBody,
            responseType: 'stream'
        };
        try {
            const res = await request.post(requestOptionsPost);
            const requestOptions = {
                url: `${this.resource}/v1.0${res.headers.location}`,
                headers: {
                    accept: 'application/json;odata.metadata=minimal'
                },
                responseType: 'json'
            };
            const teamsAsyncOperation = await request.get(requestOptions);
            if (!args.options.wait && members.length === 0) {
                await logger.log(teamsAsyncOperation);
            }
            else {
                await this.waitUntilTeamFinishedProvisioning(teamsAsyncOperation, requestOptions, logger);
                const entraGroup = await this.getEntraGroup(teamsAsyncOperation.targetResourceId, logger);
                if (members.length > 0) {
                    if (this.verbose) {
                        await logger.logToStderr('Adding members to the team...');
                    }
                    await this.addMembers(members, entraGroup.id);
                }
                await logger.log(entraGroup);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async addMembers(members, groupId) {
        for (const member of members) {
            const requestOptions = {
                url: `${this.resource}/v1.0/teams/${groupId}/members`,
                headers: {
                    'content-type': 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: member
            };
            await request.post(requestOptions);
        }
    }
    async waitUntilTeamFinishedProvisioning(teamsAsyncOperation, requestOptions, logger) {
        if (teamsAsyncOperation.status === 'succeeded') {
            if (this.verbose) {
                await logger.logToStderr('Team provisioned succesfully. Returning...');
            }
            return;
        }
        else if (teamsAsyncOperation.error) {
            throw teamsAsyncOperation.error;
        }
        else {
            if (this.verbose) {
                await logger.logToStderr(`Team still provisioning. Retrying in ${this.pollingInterval / 1000} seconds...`);
            }
            await setTimeout(this.pollingInterval);
            teamsAsyncOperation = await request.get(requestOptions);
            await this.waitUntilTeamFinishedProvisioning(teamsAsyncOperation, requestOptions, logger);
        }
    }
    async getEntraGroup(id, logger) {
        let group;
        try {
            group = await entraGroup.getGroupById(id);
        }
        catch {
            if (this.verbose) {
                await logger.logToStderr(`Error occurred on retrieving the aad group. Retrying in ${this.pollingInterval / 1000} seconds.`);
            }
            await setTimeout(this.pollingInterval);
            return await this.getEntraGroup(id, logger);
        }
        return group;
    }
    async retrieveMembersToAdd(members, role, emails, ids, userNames) {
        let itemsToProcess = [];
        if (emails) {
            itemsToProcess = await entraUser.getUserIdsByEmails(formatting.splitAndTrim(emails));
        }
        else if (ids) {
            itemsToProcess = formatting.splitAndTrim(ids);
        }
        else if (userNames) {
            itemsToProcess = await entraUser.getUserIdsByUpns(formatting.splitAndTrim(userNames));
        }
        itemsToProcess.map((item) => {
            const member = members.find((y) => y['user@odata.bind'] === `https://graph.microsoft.com/v1.0/users('${item}')`);
            if (!member) {
                members.push({
                    '@odata.type': '#microsoft.graph.aadUserConversationMember',
                    'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${item}')`,
                    roles: [role]
                });
            }
            else {
                member.roles.push(role);
            }
        });
    }
}
_TeamsTeamAddCommand_instances = new WeakSet(), _TeamsTeamAddCommand_initTelemetry = function _TeamsTeamAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            name: typeof args.options.name !== 'undefined',
            description: typeof args.options.description !== 'undefined',
            template: typeof args.options.template !== 'undefined',
            wait: !!args.options.wait,
            ownerUserNames: typeof args.options.ownerUserNames !== 'undefined',
            ownerIds: typeof args.options.ownerIds !== 'undefined',
            ownerEmails: typeof args.options.ownerEmails !== 'undefined',
            memberUserNames: typeof args.options.memberUserNames !== 'undefined',
            memberIds: typeof args.options.memberIds !== 'undefined',
            memberEmails: typeof args.options.memberEmails !== 'undefined'
        });
    });
}, _TeamsTeamAddCommand_initOptions = function _TeamsTeamAddCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name [name]'
    }, {
        option: '-d, --description [description]'
    }, {
        option: '--template [template]'
    }, {
        option: '--wait'
    }, {
        option: '--ownerUserNames [ownerUserNames]'
    }, {
        option: '--ownerIds [ownerIds]'
    }, {
        option: '--ownerEmails [ownerEmails]'
    }, {
        option: '--memberUserNames [memberUserNames]'
    }, {
        option: '--memberIds [memberIds]'
    }, {
        option: '--memberEmails [memberEmails]'
    });
}, _TeamsTeamAddCommand_initValidators = function _TeamsTeamAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.ownerUserNames) {
            const isValidUPNArrayResult = validation.isValidUserPrincipalNameArray(args.options.ownerUserNames);
            if (isValidUPNArrayResult !== true) {
                return `The following user principal names are invalid for the option 'ownerUserNames': ${isValidUPNArrayResult}.`;
            }
        }
        if (args.options.ownerEmails) {
            const isValidUPNArrayResult = validation.isValidUserPrincipalNameArray(args.options.ownerEmails);
            if (isValidUPNArrayResult !== true) {
                return `The following user principal names are invalid for the option 'ownerEmails': ${isValidUPNArrayResult}.`;
            }
        }
        if (args.options.ownerIds) {
            const isValidGUIDArrayResult = validation.isValidGuidArray(args.options.ownerIds);
            if (isValidGUIDArrayResult !== true) {
                return `The following GUIDs are invalid for the option 'ownerIds': ${isValidGUIDArrayResult}.`;
            }
        }
        if (args.options.memberUserNames) {
            const isValidUPNArrayResult = validation.isValidUserPrincipalNameArray(args.options.memberUserNames);
            if (isValidUPNArrayResult !== true) {
                return `The following user principal names are invalid for the option 'memberUserNames': ${isValidUPNArrayResult}.`;
            }
        }
        if (args.options.memberEmails) {
            const isValidUPNArrayResult = validation.isValidUserPrincipalNameArray(args.options.memberEmails);
            if (isValidUPNArrayResult !== true) {
                return `The following user principal names are invalid for the option 'memberEmails': ${isValidUPNArrayResult}.`;
            }
        }
        if (args.options.memberIds) {
            const isValidGUIDArrayResult = validation.isValidGuidArray(args.options.memberIds);
            if (isValidGUIDArrayResult !== true) {
                return `The following GUIDs are invalid for the option 'memberIds': ${isValidGUIDArrayResult}.`;
            }
        }
        return true;
    });
}, _TeamsTeamAddCommand_initOptionSets = function _TeamsTeamAddCommand_initOptionSets() {
    this.optionSets.push({
        options: ['name'],
        runsWhen: (args) => {
            return !args.options.template;
        }
    }, {
        options: ['description'],
        runsWhen: (args) => {
            return !args.options.template;
        }
    }, {
        options: ['ownerUserNames', 'ownerIds', 'ownerEmails'],
        runsWhen: (args) => {
            return args.options.ownerUserNames || args.options.ownerIds || args.options.ownerEmails;
        }
    }, {
        options: ['memberUserNames', 'memberIds', 'memberEmails'],
        runsWhen: (args) => {
            return args.options.memberUserNames || args.options.memberIds || args.options.memberEmails;
        }
    });
};
export default new TeamsTeamAddCommand();
//# sourceMappingURL=team-add.js.map