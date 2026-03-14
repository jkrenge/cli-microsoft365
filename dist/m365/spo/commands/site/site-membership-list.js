var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteMembershipListCommand_instances, _a, _SpoSiteMembershipListCommand_initTelemetry, _SpoSiteMembershipListCommand_initOptions, _SpoSiteMembershipListCommand_initValidators, _SpoSiteMembershipListCommand_initTypes;
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteMembershipListCommand extends SpoCommand {
    get name() {
        return commands.SITE_MEMBERSHIP_LIST;
    }
    get description() {
        return `Retrieves information about default site groups' membership`;
    }
    defaultProperties() {
        return ['email', 'name', 'userPrincipalName', 'associatedGroupType'];
    }
    constructor() {
        super();
        _SpoSiteMembershipListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteMembershipListCommand_instances, "m", _SpoSiteMembershipListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteMembershipListCommand_instances, "m", _SpoSiteMembershipListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteMembershipListCommand_instances, "m", _SpoSiteMembershipListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoSiteMembershipListCommand_instances, "m", _SpoSiteMembershipListCommand_initTypes).call(this);
    }
    ;
    async commandAction(logger, args) {
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.verbose);
            const roleIds = this.getRoleIds(args.options.role);
            const tenantSiteProperties = await spo.getSiteAdminPropertiesByUrl(args.options.siteUrl, false, logger, this.verbose);
            const response = await odata.getAllItems(`${spoAdminUrl}/_api/SPO.Tenant/sites/GetSiteUserGroups?siteId='${tenantSiteProperties.SiteId}'&userGroupIds=[${roleIds}]`);
            const result = args.options.output === 'json' ? this.mapResult(response, args.options.role) : this.mapListResult(response, args.options.role);
            await logger.log(result);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getRoleIds(role) {
        switch (role?.toLowerCase()) {
            case 'owner':
                return '0';
            case 'member':
                return '1';
            case 'visitor':
                return '2';
            default:
                return '0,1,2';
        }
    }
    mapResult(response, role) {
        switch (role?.toLowerCase()) {
            case 'owner':
                return { AssociatedOwnerGroup: response[0].userGroup };
            case 'member':
                return { AssociatedMemberGroup: response[0].userGroup };
            case 'visitor':
                return { AssociatedVisitorGroup: response[0].userGroup };
            default:
                return {
                    AssociatedOwnerGroup: response[0].userGroup,
                    AssociatedMemberGroup: response[1].userGroup,
                    AssociatedVisitorGroup: response[2].userGroup
                };
        }
    }
    mapListResult(response, role) {
        const mapGroup = (groupIndex, groupType) => response[groupIndex].userGroup.map(user => ({
            ...user,
            associatedGroupType: groupType
        }));
        switch (role?.toLowerCase()) {
            case 'owner':
                return mapGroup(0, 'Owner');
            case 'member':
                return mapGroup(0, 'Member');
            case 'visitor':
                return mapGroup(0, 'Visitor');
            default:
                return [
                    ...mapGroup(0, 'Owner'),
                    ...mapGroup(1, 'Member'),
                    ...mapGroup(2, 'Visitor')
                ];
        }
    }
}
_a = SpoSiteMembershipListCommand, _SpoSiteMembershipListCommand_instances = new WeakSet(), _SpoSiteMembershipListCommand_initTelemetry = function _SpoSiteMembershipListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            role: typeof args.options.role !== 'undefined'
        });
    });
}, _SpoSiteMembershipListCommand_initOptions = function _SpoSiteMembershipListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --siteUrl <siteUrl>'
    }, {
        option: '-r, --role [role]',
        autocomplete: _a.RoleNames
    });
}, _SpoSiteMembershipListCommand_initValidators = function _SpoSiteMembershipListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.role && !_a.RoleNames.some(roleName => roleName.toLocaleLowerCase() === args.options.role.toLocaleLowerCase())) {
            return `'${args.options.role}' is not a valid value for option 'role'. Valid values are: ${_a.RoleNames.join(', ')}`;
        }
        return validation.isValidSharePointUrl(args.options.siteUrl);
    });
}, _SpoSiteMembershipListCommand_initTypes = function _SpoSiteMembershipListCommand_initTypes() {
    this.types.string.push('role', 'siteUrl');
};
SpoSiteMembershipListCommand.RoleNames = ['Owner', 'Member', 'Visitor'];
export default new SpoSiteMembershipListCommand();
//# sourceMappingURL=site-membership-list.js.map