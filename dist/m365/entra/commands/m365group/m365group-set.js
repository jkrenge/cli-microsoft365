var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraM365GroupSetCommand_instances, _EntraM365GroupSetCommand_initTelemetry, _EntraM365GroupSetCommand_initOptions, _EntraM365GroupSetCommand_initOptionSets, _EntraM365GroupSetCommand_initTypes, _EntraM365GroupSetCommand_initValidators;
import { setTimeout } from 'timers/promises';
import fs from 'fs';
import path from 'path';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { accessToken } from '../../../../utils/accessToken.js';
import auth from '../../../../Auth.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
class EntraM365GroupSetCommand extends GraphCommand {
    get name() {
        return commands.M365GROUP_SET;
    }
    get description() {
        return 'Updates Microsoft 365 Group properties';
    }
    constructor() {
        super();
        _EntraM365GroupSetCommand_instances.add(this);
        this.pollingInterval = 500;
        __classPrivateFieldGet(this, _EntraM365GroupSetCommand_instances, "m", _EntraM365GroupSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupSetCommand_instances, "m", _EntraM365GroupSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupSetCommand_instances, "m", _EntraM365GroupSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupSetCommand_instances, "m", _EntraM365GroupSetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupSetCommand_instances, "m", _EntraM365GroupSetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            if ((args.options.allowExternalSenders !== undefined || args.options.autoSubscribeNewMembers !== undefined) && accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken)) {
                throw `Option 'allowExternalSenders' and 'autoSubscribeNewMembers' can only be used when using delegated permissions.`;
            }
            const groupId = args.options.id || await entraGroup.getGroupIdByDisplayName(args.options.displayName);
            const isUnifiedGroup = await entraGroup.isUnifiedGroup(groupId);
            if (!isUnifiedGroup) {
                throw Error(`Specified group with id '${groupId}' is not a Microsoft 365 group.`);
            }
            if (this.verbose) {
                await logger.logToStderr(`Updating Microsoft 365 Group ${args.options.id || args.options.displayName}...`);
            }
            if (args.options.newDisplayName || args.options.description !== undefined || args.options.isPrivate !== undefined) {
                const update = {
                    displayName: args.options.newDisplayName,
                    description: args.options.description !== '' ? args.options.description : null
                };
                if (args.options.isPrivate !== undefined) {
                    update.visibility = args.options.isPrivate ? 'Private' : 'Public';
                }
                const requestOptions = {
                    url: `${this.resource}/v1.0/groups/${groupId}`,
                    headers: {
                        'accept': 'application/json;odata.metadata=none'
                    },
                    responseType: 'json',
                    data: update
                };
                await request.patch(requestOptions);
            }
            // This has to be a separate request due to some Graph API limitations. Otherwise it will throw an error.
            if (args.options.allowExternalSenders !== undefined || args.options.autoSubscribeNewMembers !== undefined || args.options.hideFromAddressLists !== undefined || args.options.hideFromOutlookClients !== undefined) {
                const requestBody = {
                    allowExternalSenders: args.options.allowExternalSenders,
                    autoSubscribeNewMembers: args.options.autoSubscribeNewMembers,
                    hideFromAddressLists: args.options.hideFromAddressLists,
                    hideFromOutlookClients: args.options.hideFromOutlookClients
                };
                const requestOptions = {
                    url: `${this.resource}/v1.0/groups/${groupId}`,
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    },
                    responseType: 'json',
                    data: requestBody
                };
                await request.patch(requestOptions);
            }
            if (args.options.logoPath) {
                const fullPath = path.resolve(args.options.logoPath);
                if (this.verbose) {
                    await logger.logToStderr(`Setting group logo ${fullPath}...`);
                }
                const requestOptions = {
                    url: `${this.resource}/v1.0/groups/${groupId}/photo/$value`,
                    headers: {
                        'content-type': this.getImageContentType(fullPath)
                    },
                    data: fs.readFileSync(fullPath)
                };
                await this.setGroupLogo(requestOptions, EntraM365GroupSetCommand.numRepeat, logger);
            }
            else if (this.debug) {
                await logger.logToStderr('logoPath not set. Skipping');
            }
            const ownerIds = await this.getUserIds(logger, args.options.ownerIds, args.options.ownerUserNames);
            const memberIds = await this.getUserIds(logger, args.options.memberIds, args.options.memberUserNames);
            if (ownerIds.length > 0) {
                await this.updateUsers(logger, groupId, 'owners', ownerIds);
            }
            if (memberIds.length > 0) {
                await this.updateUsers(logger, groupId, 'members', memberIds);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async setGroupLogo(requestOptions, retryLeft, logger) {
        try {
            await request.put(requestOptions);
        }
        catch (err) {
            if (--retryLeft > 0) {
                await setTimeout(this.pollingInterval * (EntraM365GroupSetCommand.numRepeat - retryLeft));
                await this.setGroupLogo(requestOptions, retryLeft, logger);
            }
            else {
                throw err;
            }
        }
    }
    getImageContentType(imagePath) {
        const extension = imagePath.substring(imagePath.lastIndexOf('.')).toLowerCase();
        switch (extension) {
            case '.png':
                return 'image/png';
            case '.gif':
                return 'image/gif';
            default:
                return 'image/jpeg';
        }
    }
    async getUserIds(logger, userIds, userNames) {
        if (userIds) {
            return formatting.splitAndTrim(userIds);
        }
        if (userNames) {
            if (this.verbose) {
                await logger.logToStderr(`Retrieving user IDs...`);
            }
            return entraUser.getUserIdsByUpns(formatting.splitAndTrim(userNames));
        }
        return [];
    }
    async updateUsers(logger, groupId, role, userIds) {
        const groupUsers = await odata.getAllItems(`${this.resource}/v1.0/groups/${groupId}/${role}/microsoft.graph.user?$select=id`);
        const userIdsToAdd = userIds.filter(userId => !groupUsers.some(groupUser => groupUser.id === userId));
        const userIdsToRemove = groupUsers.filter(groupUser => !userIds.some(userId => groupUser.id === userId)).map(user => user.id);
        if (this.verbose) {
            await logger.logToStderr(`Adding ${userIdsToAdd.length} ${role}...`);
        }
        for (let i = 0; i < userIdsToAdd.length; i += 400) {
            const userIdsBatch = userIdsToAdd.slice(i, i + 400);
            const batchRequestOptions = this.getBatchRequestOptions();
            // only 20 requests per one batch are allowed
            for (let j = 0; j < userIdsBatch.length; j += 20) {
                // only 20 users can be added in one request
                const userIdsChunk = userIdsBatch.slice(j, j + 20);
                batchRequestOptions.data.requests.push({
                    id: j + 1,
                    method: 'PATCH',
                    url: `/groups/${groupId}`,
                    headers: {
                        'content-type': 'application/json;odata.metadata=none',
                        accept: 'application/json;odata.metadata=none'
                    },
                    body: {
                        [`${role}@odata.bind`]: userIdsChunk.map(u => `${this.resource}/v1.0/directoryObjects/${u}`)
                    }
                });
            }
            const res = await request.post(batchRequestOptions);
            for (const response of res.responses) {
                if (response.status !== 204) {
                    throw response.body;
                }
            }
        }
        if (this.verbose) {
            await logger.logToStderr(`Removing ${userIdsToRemove.length} ${role}...`);
        }
        for (let i = 0; i < userIdsToRemove.length; i += 20) {
            const userIdsBatch = userIdsToRemove.slice(i, i + 20);
            const batchRequestOptions = this.getBatchRequestOptions();
            userIdsBatch.map(userId => {
                batchRequestOptions.data.requests.push({
                    id: userId,
                    method: 'DELETE',
                    url: `/groups/${groupId}/${role}/${userId}/$ref`
                });
            });
            const res = await request.post(batchRequestOptions);
            for (const response of res.responses) {
                if (response.status !== 204) {
                    throw response.body;
                }
            }
        }
    }
    getBatchRequestOptions() {
        const requestOptions = {
            url: `${this.resource}/v1.0/$batch`,
            headers: {
                'content-type': 'application/json;odata.metadata=none',
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json',
            data: {
                requests: []
            }
        };
        return requestOptions;
    }
}
_EntraM365GroupSetCommand_instances = new WeakSet(), _EntraM365GroupSetCommand_initTelemetry = function _EntraM365GroupSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            displayName: typeof args.options.displayName !== 'undefined',
            newDisplayName: typeof args.options.newDisplayName !== 'undefined',
            description: typeof args.options.description !== 'undefined',
            ownerIds: typeof args.options.ownerIds !== 'undefined',
            ownerUserNames: typeof args.options.ownerUserNames !== 'undefined',
            memberIds: typeof args.options.memberIds !== 'undefined',
            memberUserNames: typeof args.options.memberUserNames !== 'undefined',
            isPrivate: !!args.options.isPrivate,
            logoPath: typeof args.options.logoPath !== 'undefined',
            allowExternalSenders: !!args.options.allowExternalSenders,
            autoSubscribeNewMembers: !!args.options.autoSubscribeNewMembers,
            hideFromAddressLists: !!args.options.hideFromAddressLists,
            hideFromOutlookClients: !!args.options.hideFromOutlookClients
        });
    });
}, _EntraM365GroupSetCommand_initOptions = function _EntraM365GroupSetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-n, --displayName [displayName]'
    }, {
        option: '--newDisplayName [newDisplayName]'
    }, {
        option: '-d, --description [description]'
    }, {
        option: '--ownerIds [ownerIds]'
    }, {
        option: '--ownerUserNames [ownerUserNames]'
    }, {
        option: '--memberIds [memberIds]'
    }, {
        option: '--memberUserNames [memberUserNames]'
    }, {
        option: '--isPrivate [isPrivate]',
        autocomplete: ['true', 'false']
    }, {
        option: '-l, --logoPath [logoPath]'
    }, {
        option: '--allowExternalSenders [allowExternalSenders]',
        autocomplete: ['true', 'false']
    }, {
        option: '--autoSubscribeNewMembers [autoSubscribeNewMembers]',
        autocomplete: ['true', 'false']
    }, {
        option: '--hideFromAddressLists [hideFromAddressLists]',
        autocomplete: ['true', 'false']
    }, {
        option: '--hideFromOutlookClients [hideFromOutlookClients]',
        autocomplete: ['true', 'false']
    });
}, _EntraM365GroupSetCommand_initOptionSets = function _EntraM365GroupSetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'displayName'] });
    this.optionSets.push({
        options: ['ownerIds', 'ownerUserNames'],
        runsWhen: (args) => {
            return args.options.ownerIds !== undefined || args.options.ownerUserNames !== undefined;
        }
    });
    this.optionSets.push({
        options: ['memberIds', 'memberUserNames'],
        runsWhen: (args) => {
            return args.options.memberIds !== undefined || args.options.memberUserNames !== undefined;
        }
    });
}, _EntraM365GroupSetCommand_initTypes = function _EntraM365GroupSetCommand_initTypes() {
    this.types.boolean.push('isPrivate', 'allowEternalSenders', 'autoSubscribeNewMembers', 'hideFromAddressLists', 'hideFromOutlookClients');
    this.types.string.push('id', 'displayName', 'newDisplayName', 'description', 'ownerIds', 'ownerUserNames', 'memberIds', 'memberUserNames', 'logoPath');
}, _EntraM365GroupSetCommand_initValidators = function _EntraM365GroupSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!args.options.newDisplayName &&
            args.options.description === undefined &&
            args.options.ownerIds === undefined &&
            args.options.ownerUserNames === undefined &&
            args.options.memberIds === undefined &&
            args.options.memberUserNames === undefined &&
            args.options.isPrivate === undefined &&
            args.options.logoPath === undefined &&
            args.options.allowExternalSenders === undefined &&
            args.options.autoSubscribeNewMembers === undefined &&
            args.options.hideFromAddressLists === undefined &&
            args.options.hideFromOutlookClients === undefined) {
            return 'Specify at least one option to update.';
        }
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        if (args.options.ownerIds) {
            const isValidGUIDArrayResult = validation.isValidGuidArray(args.options.ownerIds);
            if (isValidGUIDArrayResult !== true) {
                return `The following GUIDs are invalid for the option 'ownerIds': ${isValidGUIDArrayResult}.`;
            }
        }
        if (args.options.ownerUserNames) {
            const isValidUPNArrayResult = validation.isValidUserPrincipalNameArray(args.options.ownerUserNames);
            if (isValidUPNArrayResult !== true) {
                return `The following user principal names are invalid for the option 'ownerUserNames': ${isValidUPNArrayResult}.`;
            }
        }
        if (args.options.memberIds) {
            const isValidGUIDArrayResult = validation.isValidGuidArray(args.options.memberIds);
            if (isValidGUIDArrayResult !== true) {
                return `The following GUIDs are invalid for the option 'memberIds': ${isValidGUIDArrayResult}.`;
            }
        }
        if (args.options.memberUserNames) {
            const isValidUPNArrayResult = validation.isValidUserPrincipalNameArray(args.options.memberUserNames);
            if (isValidUPNArrayResult !== true) {
                return `The following user principal names are invalid for the option 'memberUserNames': ${isValidUPNArrayResult}.`;
            }
        }
        if (args.options.logoPath) {
            const fullPath = path.resolve(args.options.logoPath);
            if (!fs.existsSync(fullPath)) {
                return `File '${fullPath}' not found`;
            }
            if (fs.lstatSync(fullPath).isDirectory()) {
                return `Path '${fullPath}' points to a directory`;
            }
        }
        return true;
    });
};
EntraM365GroupSetCommand.numRepeat = 15;
export default new EntraM365GroupSetCommand();
//# sourceMappingURL=m365group-set.js.map