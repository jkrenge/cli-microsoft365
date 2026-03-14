var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraM365GroupAddCommand_instances, _EntraM365GroupAddCommand_initTelemetry, _EntraM365GroupAddCommand_initOptions, _EntraM365GroupAddCommand_initTypes, _EntraM365GroupAddCommand_initValidators;
import { setTimeout } from 'timers/promises';
import fs from 'fs';
import path from 'path';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraM365GroupAddCommand extends GraphCommand {
    get name() {
        return commands.M365GROUP_ADD;
    }
    get description() {
        return 'Creates a Microsoft 365 Group';
    }
    constructor() {
        super();
        _EntraM365GroupAddCommand_instances.add(this);
        this.pollingInterval = 500;
        this.allowedVisibilities = ['Private', 'Public', 'HiddenMembership'];
        __classPrivateFieldGet(this, _EntraM365GroupAddCommand_instances, "m", _EntraM365GroupAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupAddCommand_instances, "m", _EntraM365GroupAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupAddCommand_instances, "m", _EntraM365GroupAddCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupAddCommand_instances, "m", _EntraM365GroupAddCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        let group;
        const resourceBehaviorOptionsCollection = [];
        const resolvedVisibility = args.options.visibility || 'Public';
        if (this.verbose) {
            await logger.logToStderr('Creating Microsoft 365 Group...');
        }
        if (args.options.allowMembersToPost) {
            resourceBehaviorOptionsCollection.push('AllowOnlyMembersToPost');
        }
        if (args.options.hideGroupInOutlook) {
            resourceBehaviorOptionsCollection.push('HideGroupInOutlook');
        }
        if (args.options.subscribeNewGroupMembers) {
            resourceBehaviorOptionsCollection.push('SubscribeNewGroupMembers');
        }
        if (args.options.welcomeEmailDisabled) {
            resourceBehaviorOptionsCollection.push('WelcomeEmailDisabled');
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/groups`,
            headers: {
                'accept': 'application/json;odata.metadata=none'
            },
            responseType: 'json',
            data: {
                description: args.options.description,
                displayName: args.options.displayName,
                groupTypes: [
                    "Unified"
                ],
                mailEnabled: true,
                mailNickname: args.options.mailNickname,
                resourceBehaviorOptions: resourceBehaviorOptionsCollection,
                securityEnabled: false,
                visibility: resolvedVisibility
            }
        };
        try {
            const ownerIds = await this.getUserIds(logger, args.options.owners);
            const memberIds = await this.getUserIds(logger, args.options.members);
            group = await request.post(requestOptions);
            if (!args.options.logoPath) {
                if (this.debug) {
                    await logger.logToStderr('logoPath not set. Skipping');
                }
            }
            else {
                const fullPath = path.resolve(args.options.logoPath);
                if (this.verbose) {
                    await logger.logToStderr(`Setting group logo ${fullPath}...`);
                }
                const requestOptionsPhoto = {
                    url: `${this.resource}/v1.0/groups/${group.id}/photo/$value`,
                    headers: {
                        'content-type': this.getImageContentType(fullPath)
                    },
                    data: fs.readFileSync(fullPath)
                };
                await this.setGroupLogo(requestOptionsPhoto, EntraM365GroupAddCommand.numRepeat, logger);
            }
            if (ownerIds.length !== 0) {
                await Promise.all(ownerIds.map(ownerId => request.post({
                    url: `${this.resource}/v1.0/groups/${group.id}/owners/$ref`,
                    headers: {
                        'content-type': 'application/json'
                    },
                    responseType: 'json',
                    data: {
                        "@odata.id": `https://graph.microsoft.com/v1.0/users/${ownerId}`
                    }
                })));
            }
            if (memberIds.length !== 0) {
                await Promise.all(memberIds.map(memberId => request.post({
                    url: `${this.resource}/v1.0/groups/${group.id}/members/$ref`,
                    headers: {
                        'content-type': 'application/json'
                    },
                    responseType: 'json',
                    data: {
                        "@odata.id": `https://graph.microsoft.com/v1.0/users/${memberId}`
                    }
                })));
            }
            await logger.log(group);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getUserIds(logger, users) {
        if (!users) {
            if (this.debug) {
                await logger.logToStderr('No users to validate, skipping.');
            }
            return [];
        }
        if (this.verbose) {
            await logger.logToStderr('Retrieving user information.');
        }
        const userArr = users.split(',').map(o => o.trim());
        const promises = userArr.map(user => {
            const requestOptions = {
                url: `${this.resource}/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(user)}'&$select=id,userPrincipalName`,
                headers: {
                    'content-type': 'application/json'
                },
                responseType: 'json'
            };
            return request.get(requestOptions);
        });
        const usersRes = await Promise.all(promises);
        let userUpns = [];
        userUpns = usersRes.map(res => res.value[0]?.userPrincipalName);
        const userIds = usersRes.map(res => res.value[0]?.id);
        // Find the members where no graph response was found
        const invalidUsers = userArr.filter(user => !userUpns.some((upn) => upn?.toLowerCase() === user.toLowerCase()));
        if (invalidUsers && invalidUsers.length > 0) {
            throw `Cannot proceed with group creation. The following users provided are invalid : ${invalidUsers.join(',')}`;
        }
        return userIds;
    }
    async setGroupLogo(requestOptions, retryLeft, logger) {
        try {
            await request.put(requestOptions);
        }
        catch (err) {
            if (--retryLeft > 0) {
                await setTimeout(this.pollingInterval * (EntraM365GroupAddCommand.numRepeat - retryLeft));
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
}
_EntraM365GroupAddCommand_instances = new WeakSet(), _EntraM365GroupAddCommand_initTelemetry = function _EntraM365GroupAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            description: typeof args.options.description !== 'undefined',
            owners: typeof args.options.owners !== 'undefined',
            members: typeof args.options.members !== 'undefined',
            logoPath: typeof args.options.logoPath !== 'undefined',
            visibility: typeof args.options.visibility !== 'undefined',
            allowMembersToPost: !!args.options.allowMembersToPost,
            hideGroupInOutlook: !!args.options.hideGroupInOutlook,
            subscribeNewGroupMembers: !!args.options.subscribeNewGroupMembers,
            welcomeEmailDisabled: !!args.options.welcomeEmailDisabled
        });
    });
}, _EntraM365GroupAddCommand_initOptions = function _EntraM365GroupAddCommand_initOptions() {
    this.options.unshift({
        option: '-n, --displayName <displayName>'
    }, {
        option: '-m, --mailNickname <mailNickname>'
    }, {
        option: '-d, --description [description]'
    }, {
        option: '--owners [owners]'
    }, {
        option: '--members [members]'
    }, {
        option: '--visibility [visibility]',
        autocomplete: this.allowedVisibilities
    }, {
        option: '--allowMembersToPost [allowMembersToPost]',
        autocomplete: ['true', 'false']
    }, {
        option: '--hideGroupInOutlook [hideGroupInOutlook]',
        autocomplete: ['true', 'false']
    }, {
        option: '--subscribeNewGroupMembers [subscribeNewGroupMembers]',
        autocomplete: ['true', 'false']
    }, {
        option: '--welcomeEmailDisabled [welcomeEmailDisabled]',
        autocomplete: ['true', 'false']
    }, {
        option: '-l, --logoPath [logoPath]'
    });
}, _EntraM365GroupAddCommand_initTypes = function _EntraM365GroupAddCommand_initTypes() {
    this.types.string.push('displayName', 'mailNickname', 'description', 'owners', 'members', 'visibility', 'logoPath');
    this.types.boolean.push('allowMembersToPost', 'hideGroupInOutlook', 'subscribeNewGroupMembers', 'welcomeEmailDisabled');
}, _EntraM365GroupAddCommand_initValidators = function _EntraM365GroupAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.owners) {
            const owners = args.options.owners.split(',').map(o => o.trim());
            for (let i = 0; i < owners.length; i++) {
                if (owners[i].indexOf('@') < 0) {
                    return `${owners[i]} is not a valid userPrincipalName`;
                }
            }
        }
        if (args.options.members) {
            const members = args.options.members.split(',').map(m => m.trim());
            for (let i = 0; i < members.length; i++) {
                if (members[i].indexOf('@') < 0) {
                    return `${members[i]} is not a valid userPrincipalName`;
                }
            }
        }
        if (args.options.mailNickname.indexOf(' ') !== -1) {
            return 'The option mailNickname cannot contain spaces.';
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
        if (args.options.visibility && this.allowedVisibilities.map(x => x.toLowerCase()).indexOf(args.options.visibility.toLowerCase()) === -1) {
            return `${args.options.visibility} is not a valid visibility. Allowed values are ${this.allowedVisibilities.join(', ')}`;
        }
        return true;
    });
};
EntraM365GroupAddCommand.numRepeat = 15;
export default new EntraM365GroupAddCommand();
//# sourceMappingURL=m365group-add.js.map