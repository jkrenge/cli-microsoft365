var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteAdminRemoveCommand_instances, _SpoSiteAdminRemoveCommand_initTelemetry, _SpoSiteAdminRemoveCommand_initOptions, _SpoSiteAdminRemoveCommand_initValidators, _SpoSiteAdminRemoveCommand_initOptionSets, _SpoSiteAdminRemoveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteAdminRemoveCommand extends SpoCommand {
    get name() {
        return commands.SITE_ADMIN_REMOVE;
    }
    get description() {
        return 'Removes a user or group as site collection administrator';
    }
    constructor() {
        super();
        _SpoSiteAdminRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteAdminRemoveCommand_instances, "m", _SpoSiteAdminRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteAdminRemoveCommand_instances, "m", _SpoSiteAdminRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteAdminRemoveCommand_instances, "m", _SpoSiteAdminRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoSiteAdminRemoveCommand_instances, "m", _SpoSiteAdminRemoveCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoSiteAdminRemoveCommand_instances, "m", _SpoSiteAdminRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (!args.options.force) {
                const principalToDelete = args.options.groupId || args.options.groupName ? 'group' : 'user';
                const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove specified ${principalToDelete} from the site administrators list ${args.options.siteUrl}?` });
                if (!result) {
                    return;
                }
            }
            const loginNameToRemove = await this.getCorrectLoginName(args.options);
            if (args.options.asAdmin) {
                await this.callActionAsAdmin(logger, args, loginNameToRemove);
                return;
            }
            await this.callAction(logger, args, loginNameToRemove);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async callActionAsAdmin(logger, args, loginNameToRemove) {
        if (this.verbose) {
            await logger.logToStderr('Removing site administrator as an administrator...');
        }
        const adminUrl = await spo.getSpoAdminUrl(logger, this.debug);
        const tenantSiteProperties = await spo.getSiteAdminPropertiesByUrl(args.options.siteUrl, false, logger, this.verbose);
        const primaryAdminLoginName = await spo.getPrimaryAdminLoginNameAsAdmin(adminUrl, tenantSiteProperties.SiteId, logger, this.verbose);
        if (loginNameToRemove === primaryAdminLoginName) {
            throw 'You cannot remove the primary site collection administrator.';
        }
        const existingAdmins = await this.getSiteAdmins(adminUrl, tenantSiteProperties.SiteId);
        const adminsToSet = existingAdmins.filter(u => u.loginName.toLowerCase() !== loginNameToRemove.toLowerCase());
        await this.setSiteAdminsAsAdmin(adminUrl, tenantSiteProperties.SiteId, adminsToSet);
    }
    async getSiteAdmins(adminUrl, siteId) {
        const requestOptions = {
            url: `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${siteId}'`,
            headers: {
                accept: 'application/json;odata=nometadata',
                'content-type': 'application/json;charset=utf-8'
            },
            responseType: 'json'
        };
        const response = await request.post(requestOptions);
        return response.value;
    }
    async getCorrectLoginName(options) {
        if (options.userId || options.userName) {
            const userPrincipalName = options.userName ? options.userName : await entraUser.getUpnByUserId(options.userId);
            if (userPrincipalName) {
                return `i:0#.f|membership|${userPrincipalName}`;
            }
            throw 'User not found.';
        }
        else {
            const group = options.groupId ? await entraGroup.getGroupById(options.groupId) : await entraGroup.getGroupByDisplayName(options.groupName);
            //for entra groups, M365 groups have an associated email and security groups don't
            if (group?.mail) {
                //M365 group is prefixed with c:0o.c|federateddirectoryclaimprovider
                return `c:0o.c|federateddirectoryclaimprovider|${group.id}`;
            }
            else {
                //security group is prefixed with c:0t.c|tenant
                return `c:0t.c|tenant|${group?.id}`;
            }
        }
    }
    async setSiteAdminsAsAdmin(adminUrl, siteId, admins) {
        const requestOptions = {
            url: `${adminUrl}/_api/SPOInternalUseOnly.Tenant/SetSiteSecondaryAdministrators`,
            headers: {
                accept: 'application/json;odata=nometadata',
                'content-type': 'application/json;charset=utf-8'
            },
            data: {
                secondaryAdministratorsFieldsData: {
                    siteId: siteId,
                    secondaryAdministratorLoginNames: admins.map(u => u.loginName)
                }
            },
            responseType: 'json'
        };
        await request.post(requestOptions);
    }
    async callAction(logger, args, loginNameToRemove) {
        if (this.verbose) {
            await logger.logToStderr('Removing site administrator...');
        }
        const primaryOwnerLogin = await spo.getPrimaryOwnerLoginFromSite(args.options.siteUrl, logger, this.verbose);
        if (loginNameToRemove === primaryOwnerLogin) {
            throw 'You cannot remove the primary site collection administrator.';
        }
        const requestOptions = {
            url: `${args.options.siteUrl}/_api/web/siteusers('${formatting.encodeQueryParameter(loginNameToRemove)}')`,
            headers: {
                'accept': 'application/json',
                'X-Http-Method': 'MERGE',
                'If-Match': '*'
            },
            data: { IsSiteAdmin: false },
            responseType: 'json'
        };
        await request.post(requestOptions);
    }
}
_SpoSiteAdminRemoveCommand_instances = new WeakSet(), _SpoSiteAdminRemoveCommand_initTelemetry = function _SpoSiteAdminRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            groupId: typeof args.options.groupId !== 'undefined',
            groupName: typeof args.options.groupName !== 'undefined',
            force: !!args.options.force,
            asAdmin: !!args.options.asAdmin
        });
    });
}, _SpoSiteAdminRemoveCommand_initOptions = function _SpoSiteAdminRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --siteUrl <siteUrl>'
    }, {
        option: '--userId [userId]'
    }, {
        option: '--userName [userName]'
    }, {
        option: '--groupId [groupId]'
    }, {
        option: '--groupName [groupName]'
    }, {
        option: '--asAdmin'
    }, {
        option: '-f, --force'
    });
}, _SpoSiteAdminRemoveCommand_initValidators = function _SpoSiteAdminRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.userId &&
            !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid userName`;
        }
        if (args.options.groupId &&
            !validation.isValidGuid(args.options.groupId)) {
            return `${args.options.groupId} is not a valid GUID`;
        }
        return validation.isValidSharePointUrl(args.options.siteUrl);
    });
}, _SpoSiteAdminRemoveCommand_initOptionSets = function _SpoSiteAdminRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['userId', 'userName', 'groupId', 'groupName'] });
}, _SpoSiteAdminRemoveCommand_initTypes = function _SpoSiteAdminRemoveCommand_initTypes() {
    this.types.string.push('siteUrl', 'userId', 'userName', 'groupId', 'groupName');
    this.types.boolean.push('force', 'asAdmin');
};
export default new SpoSiteAdminRemoveCommand();
//# sourceMappingURL=site-admin-remove.js.map