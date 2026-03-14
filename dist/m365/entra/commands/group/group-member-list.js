var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraGroupMemberListCommand_instances, _EntraGroupMemberListCommand_initTelemetry, _EntraGroupMemberListCommand_initOptions, _EntraGroupMemberListCommand_initOptionSets, _EntraGroupMemberListCommand_initValidators;
import { entraGroup } from '../../../../utils/entraGroup.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraGroupMemberListCommand extends GraphCommand {
    get name() {
        return commands.GROUP_MEMBER_LIST;
    }
    get description() {
        return 'Lists members of a specific Entra group';
    }
    defaultProperties() {
        return ['id', 'displayName', 'userPrincipalName', 'roles'];
    }
    constructor() {
        super();
        _EntraGroupMemberListCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraGroupMemberListCommand_instances, "m", _EntraGroupMemberListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraGroupMemberListCommand_instances, "m", _EntraGroupMemberListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraGroupMemberListCommand_instances, "m", _EntraGroupMemberListCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _EntraGroupMemberListCommand_instances, "m", _EntraGroupMemberListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const groupId = await this.getGroupId(args.options, logger);
            const users = [];
            if (!args.options.role || args.options.role === 'Owner') {
                const owners = await this.getUsers(args.options, 'Owners', groupId, logger);
                owners.forEach(owner => users.push({ ...owner, roles: ['Owner'] }));
            }
            if (!args.options.role || args.options.role === 'Member') {
                const members = await this.getUsers(args.options, 'Members', groupId, logger);
                members.forEach((member) => {
                    const user = users.find((u) => u.id === member.id);
                    if (user !== undefined) {
                        user.roles.push('Member');
                    }
                    else {
                        users.push({ ...member, roles: ['Member'] });
                    }
                });
            }
            await logger.log(users);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getGroupId(options, logger) {
        if (options.groupId) {
            return options.groupId;
        }
        if (this.verbose) {
            await logger.logToStderr('Retrieving Group Id...');
        }
        return await entraGroup.getGroupIdByDisplayName(options.groupName);
    }
    async getUsers(options, role, groupId, logger) {
        const { properties, filter } = options;
        if (this.verbose) {
            await logger.logToStderr(`Retrieving ${role} of the group with id ${groupId}`);
        }
        const selectProperties = properties ?
            `${properties.split(',').filter(f => f.toLowerCase() !== 'id').concat('id').map(p => p.trim()).join(',')}` :
            'id,displayName,userPrincipalName,givenName,surname';
        const allSelectProperties = selectProperties.split(',');
        const propertiesWithSlash = allSelectProperties.filter(item => item.includes('/'));
        let fieldExpand = '';
        propertiesWithSlash.forEach(p => {
            if (fieldExpand.length > 0) {
                fieldExpand += ',';
            }
            fieldExpand += `${p.split('/')[0]}($select=${p.split('/')[1]})`;
        });
        const expandParam = fieldExpand.length > 0 ? `&$expand=${fieldExpand}` : '';
        const selectParam = allSelectProperties.filter(item => !item.includes('/'));
        const endpoint = `${this.resource}/v1.0/groups/${groupId}/${role}?$select=${selectParam}${expandParam}`;
        let users;
        if (filter) {
            // While using the filter, we need to specify the ConsistencyLevel header.
            // Can be refactored when the header is no longer necessary.
            const requestOptions = {
                url: `${endpoint}&$filter=${encodeURIComponent(filter)}&$count=true`,
                headers: {
                    accept: 'application/json;odata.metadata=none',
                    ConsistencyLevel: 'eventual'
                },
                responseType: 'json'
            };
            users = await odata.getAllItems(requestOptions);
        }
        else {
            users = await odata.getAllItems(endpoint);
        }
        return users;
    }
}
_EntraGroupMemberListCommand_instances = new WeakSet(), _EntraGroupMemberListCommand_initTelemetry = function _EntraGroupMemberListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            groupId: typeof args.options.groupId !== 'undefined',
            groupName: typeof args.options.groupName !== 'undefined',
            role: typeof args.options.role !== 'undefined',
            properties: typeof args.options.properties !== 'undefined',
            filter: typeof args.options.filter !== 'undefined'
        });
    });
}, _EntraGroupMemberListCommand_initOptions = function _EntraGroupMemberListCommand_initOptions() {
    this.options.unshift({
        option: "-i, --groupId [groupId]"
    }, {
        option: "-n, --groupName [groupName]"
    }, {
        option: "-r, --role [role]",
        autocomplete: ["Owner", "Member"]
    }, {
        option: "-p, --properties [properties]"
    }, {
        option: "-f, --filter [filter]"
    });
}, _EntraGroupMemberListCommand_initOptionSets = function _EntraGroupMemberListCommand_initOptionSets() {
    this.optionSets.push({
        options: ['groupId', 'groupName']
    });
}, _EntraGroupMemberListCommand_initValidators = function _EntraGroupMemberListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.groupId && !validation.isValidGuid(args.options.groupId)) {
            return `${args.options.groupId} is not a valid GUID`;
        }
        if (args.options.role) {
            if (['Owner', 'Member'].indexOf(args.options.role) === -1) {
                return `${args.options.role} is not a valid role value. Allowed values Owner|Member`;
            }
        }
        return true;
    });
};
export default new EntraGroupMemberListCommand();
//# sourceMappingURL=group-member-list.js.map