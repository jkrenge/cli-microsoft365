var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraGroupAddCommand_instances, _EntraGroupAddCommand_initOptions, _EntraGroupAddCommand_initValidators, _EntraGroupAddCommand_initTelemetry;
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import request from '../../../../request.js';
import { entraUser } from '../../../../utils/entraUser.js';
class EntraGroupAddCommand extends GraphCommand {
    get name() {
        return commands.GROUP_ADD;
    }
    get description() {
        return 'Creates a Microsoft Entra group';
    }
    allowUnknownOptions() {
        return true;
    }
    constructor() {
        super();
        _EntraGroupAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraGroupAddCommand_instances, "m", _EntraGroupAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraGroupAddCommand_instances, "m", _EntraGroupAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraGroupAddCommand_instances, "m", _EntraGroupAddCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        let group;
        try {
            const manifest = this.createRequestBody(args.options);
            const requestOptions = {
                url: `${this.resource}/v1.0/groups`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: manifest
            };
            const ownerIds = await this.getUserIds(logger, args.options.ownerIds, args.options.ownerUserNames);
            const memberIds = await this.getUserIds(logger, args.options.memberIds, args.options.memberUserNames);
            group = await request.post(requestOptions);
            if (ownerIds.length !== 0) {
                await this.addUsers(group.id, 'owners', ownerIds);
            }
            if (memberIds.length !== 0) {
                await this.addUsers(group.id, 'members', memberIds);
            }
            await logger.log(group);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    ;
    createRequestBody(options) {
        const requestBody = {
            displayName: options.displayName,
            description: options.description,
            mailNickName: options.mailNickname ?? this.generateMailNickname(),
            visibility: options.visibility ?? 'Public',
            groupTypes: options.type === 'microsoft365' ? ['Unified'] : [],
            mailEnabled: options.type === 'security' ? false : true,
            securityEnabled: true
        };
        this.addUnknownOptionsToPayload(requestBody, options);
        return requestBody;
    }
    generateMailNickname() {
        return `Group${Math.floor(Math.random() * 1000000)}`;
    }
    async getUserIds(logger, userIds, userNames) {
        if (userIds) {
            return userIds.split(',').map(o => o.trim());
        }
        if (!userNames) {
            if (this.verbose) {
                await logger.logToStderr('No users to validate, skipping.');
            }
            return [];
        }
        if (this.verbose) {
            await logger.logToStderr('Retrieving user information.');
        }
        const userArr = userNames.split(',').map(o => o.trim());
        if (this.verbose) {
            await logger.logToStderr('Retrieving ID(s) of user(s)...');
        }
        return entraUser.getUserIdsByUpns(userArr);
    }
    async addUsers(groupId, role, userIds) {
        for (let i = 0; i < userIds.length; i += 400) {
            const userIdsBatch = userIds.slice(i, i + 400);
            const requestOptions = {
                url: `${this.resource}/v1.0/$batch`,
                headers: {
                    'content-type': 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: {
                    requests: []
                }
            };
            // only 20 requests per one batch are allowed
            for (let j = 0; j < userIdsBatch.length; j += 20) {
                // only 20 users can be added in one request
                const userIdsChunk = userIdsBatch.slice(j, j + 20);
                requestOptions.data.requests.push({
                    id: j + 1,
                    method: 'PATCH',
                    url: `/groups/${groupId}`,
                    headers: {
                        'content-type': 'application/json;odata.metadata=none'
                    },
                    body: {
                        [`${role}@odata.bind`]: userIdsChunk.map(u => `${this.resource}/v1.0/directoryObjects/${u}`)
                    }
                });
            }
            const res = await request.post(requestOptions);
            for (const response of res.responses) {
                if (response.status !== 204) {
                    throw response.body;
                }
            }
        }
    }
}
_EntraGroupAddCommand_instances = new WeakSet(), _EntraGroupAddCommand_initOptions = function _EntraGroupAddCommand_initOptions() {
    this.options.unshift({
        option: '-n, --displayName <displayName>'
    }, {
        option: '-d, --description [description]'
    }, {
        option: '-t, --type <type>',
        autocomplete: ['microsoft365', 'security']
    }, {
        option: '-m, --mailNickname [mailNickname]'
    }, {
        option: '--ownerIds [ownerIds]'
    }, {
        option: '--ownerUserNames [ownerUserNames]'
    }, {
        option: '--memberIds [memberIds]'
    }, {
        option: '--memberUserNames [memberUserNames]'
    }, {
        option: '--visibility [visibility]',
        autocomplete: ['Public', 'Private', 'HiddenMembership']
    });
}, _EntraGroupAddCommand_initValidators = function _EntraGroupAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.displayName.length > 256) {
            return `The maximum amount of characters for 'displayName' is 256.`;
        }
        if (args.options.mailNickname) {
            if (!validation.isValidMailNickname(args.options.mailNickname)) {
                return `Value for option 'mailNickname' must contain only characters in the ASCII character set 0-127 except the following: @ () \\ [] " ; : <> , SPACE.`;
            }
            if (args.options.mailNickname.length > 64) {
                return `The maximum amount of characters for 'mailNickname' is 64.`;
            }
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
        if (['microsoft365', 'security'].indexOf(args.options.type) === -1) {
            return `Option 'type' must be one of the following values: microsoft365, security.`;
        }
        if (args.options.type === 'microsoft365' && !args.options.visibility) {
            return `Option 'visibility' must be specified if the option 'type' is set to microsoft365`;
        }
        if (args.options.visibility && ['Public', 'Private', 'HiddenMembership'].indexOf(args.options.visibility) === -1) {
            return `Option 'visibility' must be one of the following values: Public, Private, HiddenMembership.`;
        }
        return true;
    });
}, _EntraGroupAddCommand_initTelemetry = function _EntraGroupAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            displayName: typeof args.options.displayName !== 'undefined',
            description: typeof args.options.description !== 'undefined',
            type: typeof args.options.type !== 'undefined',
            mailNickname: typeof args.options.mailNickname !== 'undefined',
            ownerIds: typeof args.options.ownerIds !== 'undefined',
            ownerUserNames: typeof args.options.ownerUserNames !== 'undefined',
            memberIds: typeof args.options.memberIds !== 'undefined',
            memberUserNames: typeof args.options.memberUserNames !== 'undefined',
            visibility: typeof args.options.visibility !== 'undefined'
        });
        this.trackUnknownOptions(this.telemetryProperties, args.options);
    });
};
export default new EntraGroupAddCommand();
//# sourceMappingURL=group-add.js.map