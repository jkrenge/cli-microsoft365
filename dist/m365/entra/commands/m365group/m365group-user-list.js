var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraM365GroupUserListCommand_instances, _EntraM365GroupUserListCommand_initTelemetry, _EntraM365GroupUserListCommand_initOptions, _EntraM365GroupUserListCommand_initOptionSets, _EntraM365GroupUserListCommand_initValidators;
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
class EntraM365GroupUserListCommand extends GraphCommand {
    get name() {
        return commands.M365GROUP_USER_LIST;
    }
    get description() {
        return "Lists users for the specified Microsoft 365 group";
    }
    constructor() {
        super();
        _EntraM365GroupUserListCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraM365GroupUserListCommand_instances, "m", _EntraM365GroupUserListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupUserListCommand_instances, "m", _EntraM365GroupUserListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupUserListCommand_instances, "m", _EntraM365GroupUserListCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupUserListCommand_instances, "m", _EntraM365GroupUserListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const groupId = await this.getGroupId(args.options, logger);
            const isUnifiedGroup = await entraGroup.isUnifiedGroup(groupId);
            if (!isUnifiedGroup) {
                throw Error(`Specified group '${args.options.groupId || args.options.groupDisplayName}' is not a Microsoft 365 group.`);
            }
            let users = [];
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
            if (args.options.role) {
                users = users.filter(i => i.roles.indexOf(args.options.role) > -1);
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
        return await entraGroup.getGroupIdByDisplayName(options.groupDisplayName);
    }
    async getUsers(options, role, groupId, logger) {
        const { properties, filter } = options;
        if (this.verbose) {
            await logger.logToStderr(`Retrieving ${role} of the group with id ${groupId}`);
        }
        const selectProperties = properties ?
            `${properties.split(',').filter(f => f.toLowerCase() !== 'id').concat('id').map(p => p.trim()).join(',')}` :
            'id,displayName,userPrincipalName,givenName,surname,userType';
        const allSelectProperties = selectProperties.split(',');
        const propertiesWithSlash = allSelectProperties.filter(item => item.includes('/'));
        const fieldsToExpand = [];
        propertiesWithSlash.forEach(p => {
            const propertiesSplit = p.split('/');
            fieldsToExpand.push(`${propertiesSplit[0]}($select=${propertiesSplit[1]})`);
        });
        const fieldExpand = fieldsToExpand.join(',');
        const expandParam = fieldExpand.length > 0 ? `&$expand=${fieldExpand}` : '';
        const selectParam = allSelectProperties.filter(item => !item.includes('/'));
        const endpoint = `${this.resource}/v1.0/groups/${groupId}/${role}/microsoft.graph.user?$select=${selectParam}${expandParam}`;
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
            return await odata.getAllItems(requestOptions);
        }
        else {
            return await odata.getAllItems(endpoint);
        }
    }
}
_EntraM365GroupUserListCommand_instances = new WeakSet(), _EntraM365GroupUserListCommand_initTelemetry = function _EntraM365GroupUserListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            groupId: typeof args.options.groupId !== 'undefined',
            groupDisplayName: typeof args.options.groupDisplayName !== 'undefined',
            role: typeof args.options.role !== 'undefined',
            properties: typeof args.options.properties !== 'undefined',
            filter: typeof args.options.filter !== 'undefined'
        });
    });
}, _EntraM365GroupUserListCommand_initOptions = function _EntraM365GroupUserListCommand_initOptions() {
    this.options.unshift({
        option: "-i, --groupId [groupId]"
    }, {
        option: "-n, --groupDisplayName [groupDisplayName]"
    }, {
        option: "-r, --role [type]",
        autocomplete: ["Owner", "Member"]
    }, {
        option: "-p, --properties [properties]"
    }, {
        option: "-f, --filter [filter]"
    });
}, _EntraM365GroupUserListCommand_initOptionSets = function _EntraM365GroupUserListCommand_initOptionSets() {
    this.optionSets.push({
        options: ['groupId', 'groupDisplayName']
    });
}, _EntraM365GroupUserListCommand_initValidators = function _EntraM365GroupUserListCommand_initValidators() {
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
export default new EntraM365GroupUserListCommand();
//# sourceMappingURL=m365group-user-list.js.map