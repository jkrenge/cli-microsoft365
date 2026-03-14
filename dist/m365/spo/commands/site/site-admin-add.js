var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteAdminAddCommand_instances, _SpoSiteAdminAddCommand_initTelemetry, _SpoSiteAdminAddCommand_initOptions, _SpoSiteAdminAddCommand_initValidators, _SpoSiteAdminAddCommand_initOptionSets, _SpoSiteAdminAddCommand_initTypes;
import config from '../../../../config.js';
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteAdminAddCommand extends SpoCommand {
    get name() {
        return commands.SITE_ADMIN_ADD;
    }
    get description() {
        return 'Adds a user or group as a site collection administrator';
    }
    constructor() {
        super();
        _SpoSiteAdminAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteAdminAddCommand_instances, "m", _SpoSiteAdminAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteAdminAddCommand_instances, "m", _SpoSiteAdminAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteAdminAddCommand_instances, "m", _SpoSiteAdminAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoSiteAdminAddCommand_instances, "m", _SpoSiteAdminAddCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoSiteAdminAddCommand_instances, "m", _SpoSiteAdminAddCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            const loginNameToAdd = await this.getCorrectLoginName(args.options);
            if (args.options.asAdmin) {
                await this.callActionAsAdmin(logger, args, loginNameToAdd);
                return;
            }
            await this.callAction(logger, args, loginNameToAdd);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async callActionAsAdmin(logger, args, loginNameToAdd) {
        if (this.verbose) {
            await logger.logToStderr('Adding site administrator as an administrator...');
        }
        const adminUrl = await spo.getSpoAdminUrl(logger, this.debug);
        const tenantSiteProperties = await spo.getSiteAdminPropertiesByUrl(args.options.siteUrl, false, logger, this.verbose);
        const siteAdmins = (await this.getSiteAdmins(adminUrl, tenantSiteProperties.SiteId)).map(u => u.loginName);
        siteAdmins.push(loginNameToAdd);
        await this.setSiteAdminsAsAdmin(adminUrl, tenantSiteProperties.SiteId, siteAdmins);
        if (args.options.primary) {
            await this.setPrimaryAdminAsAdmin(adminUrl, tenantSiteProperties.SiteId, loginNameToAdd);
        }
    }
    async getSiteAdmins(adminUrl, siteId) {
        const requestOptions = {
            url: `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${siteId}'`,
            headers: {
                accept: 'application/json;odata=nometadata',
                'content-type': 'application/json;charset=utf-8'
            }
        };
        const response = await request.post(requestOptions);
        const responseContent = JSON.parse(response);
        return responseContent.value;
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
                    secondaryAdministratorLoginNames: admins
                }
            }
        };
        return request.post(requestOptions);
    }
    async setPrimaryAdminAsAdmin(adminUrl, siteId, adminLogin) {
        const requestOptions = {
            url: `${adminUrl}/_api/SPO.Tenant/sites('${siteId}')`,
            headers: {
                accept: 'application/json;odata=nometadata',
                'content-type': 'application/json;charset=utf-8'
            },
            data: {
                Owner: adminLogin,
                SetOwnerWithoutUpdatingSecondaryAdmin: true
            }
        };
        return request.patch(requestOptions);
    }
    async callAction(logger, args, loginNameToAdd) {
        if (this.verbose) {
            await logger.logToStderr('Adding site administrator...');
        }
        const ensuredUserData = await this.ensureUser(args, loginNameToAdd);
        await this.setSiteAdmin(args.options.siteUrl, loginNameToAdd);
        if (args.options.primary) {
            const siteId = await spo.getSiteIdBySPApi(args.options.siteUrl, logger, this.verbose);
            const previousPrimaryOwner = await this.getSiteOwnerLoginName(args.options.siteUrl);
            await this.setPrimaryOwnerLoginFromSite(logger, args.options.siteUrl, siteId, ensuredUserData);
            await this.setSiteAdmin(args.options.siteUrl, previousPrimaryOwner);
        }
    }
    async ensureUser(args, loginName) {
        const requestOptions = {
            url: `${args.options.siteUrl}/_api/web/ensureuser`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            data: {
                logonName: loginName
            },
            responseType: 'json'
        };
        return request.post(requestOptions);
    }
    async setSiteAdmin(siteUrl, loginName) {
        const requestOptions = {
            url: `${siteUrl}/_api/web/siteusers('${formatting.encodeQueryParameter(loginName)}')`,
            headers: {
                'accept': 'application/json',
                'X-Http-Method': 'MERGE',
                'If-Match': '*'
            },
            data: { IsSiteAdmin: true },
            responseType: 'json'
        };
        return request.post(requestOptions);
    }
    async getSiteOwnerLoginName(siteUrl) {
        const requestOptions = {
            url: `${siteUrl}/_api/site/owner?$select=LoginName`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        return response.LoginName;
    }
    async setPrimaryOwnerLoginFromSite(logger, siteUrl, siteId, loginName) {
        const res = await spo.ensureFormDigest(siteUrl, logger, undefined, this.debug);
        const body = `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}"><Actions><SetProperty Id="10" ObjectPathId="2" Name="Owner"><Parameter ObjectPathId="3" /></SetProperty></Actions><ObjectPaths><Property Id="2" ParentId="0" Name="Site" /><Identity Id="3" Name="6d452ba1-40a8-8000-e00d-46e1adaa12bf|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${siteId}:u:${loginName.Id}" /><StaticProperty Id="0" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`;
        const requestOptions = {
            url: `${siteUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
                'X-RequestDigest': res.FormDigestValue
            },
            data: body
        };
        return request.post(requestOptions);
    }
}
_SpoSiteAdminAddCommand_instances = new WeakSet(), _SpoSiteAdminAddCommand_initTelemetry = function _SpoSiteAdminAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            groupId: typeof args.options.groupId !== 'undefined',
            groupName: typeof args.options.groupName !== 'undefined',
            primary: !!args.options.primary,
            asAdmin: !!args.options.asAdmin
        });
    });
}, _SpoSiteAdminAddCommand_initOptions = function _SpoSiteAdminAddCommand_initOptions() {
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
        option: '--primary'
    }, {
        option: '--asAdmin'
    });
}, _SpoSiteAdminAddCommand_initValidators = function _SpoSiteAdminAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.userId &&
            !validation.isValidGuid(args.options.userId)) {
            return `'${args.options.userId}' is not a valid GUID for option 'userId'`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `'${args.options.userName}' is not a valid 'userName'`;
        }
        if (args.options.groupId &&
            !validation.isValidGuid(args.options.groupId)) {
            return `'${args.options.groupId}' is not a valid GUID for option 'groupId'`;
        }
        return validation.isValidSharePointUrl(args.options.siteUrl);
    });
}, _SpoSiteAdminAddCommand_initOptionSets = function _SpoSiteAdminAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['userId', 'userName', 'groupId', 'groupName'] });
}, _SpoSiteAdminAddCommand_initTypes = function _SpoSiteAdminAddCommand_initTypes() {
    this.types.string.push('siteUrl', 'userId', 'userName', 'groupId', 'groupName');
    this.types.boolean.push('primary', 'asAdmin');
};
export default new SpoSiteAdminAddCommand();
//# sourceMappingURL=site-admin-add.js.map