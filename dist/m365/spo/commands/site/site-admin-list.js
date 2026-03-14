var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteAdminListCommand_instances, _SpoSiteAdminListCommand_initTelemetry, _SpoSiteAdminListCommand_initOptions, _SpoSiteAdminListCommand_initValidators, _SpoSiteAdminListCommand_initTypes;
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { ListPrincipalType } from '../list/ListPrincipalType.js';
class SpoSiteAdminListCommand extends SpoCommand {
    get name() {
        return commands.SITE_ADMIN_LIST;
    }
    get description() {
        return 'Lists all administrators of a specific SharePoint site';
    }
    defaultProperties() {
        return ['Id', 'LoginName', 'Title', 'PrincipalTypeString'];
    }
    constructor() {
        super();
        _SpoSiteAdminListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteAdminListCommand_instances, "m", _SpoSiteAdminListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteAdminListCommand_instances, "m", _SpoSiteAdminListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteAdminListCommand_instances, "m", _SpoSiteAdminListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoSiteAdminListCommand_instances, "m", _SpoSiteAdminListCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (args.options.asAdmin) {
                await this.callActionAsAdmin(logger, args);
                return;
            }
            await this.callAction(logger, args);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async callActionAsAdmin(logger, args) {
        if (this.verbose) {
            await logger.logToStderr('Retrieving site administrators as an administrator...');
        }
        const adminUrl = await spo.getSpoAdminUrl(logger, this.debug);
        const tenantSiteProperties = await spo.getSiteAdminPropertiesByUrl(args.options.siteUrl, false, logger, this.verbose);
        const requestOptions = {
            url: `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${tenantSiteProperties.SiteId}'`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const response = await request.post(requestOptions);
        const primaryAdminLoginName = await spo.getPrimaryAdminLoginNameAsAdmin(adminUrl, tenantSiteProperties.SiteId, logger, this.verbose);
        const mappedResult = response.value.map((u) => ({
            Email: u.email,
            LoginName: u.loginName,
            Title: u.name,
            IsPrimaryAdmin: u.loginName === primaryAdminLoginName
        }));
        await logger.log(mappedResult);
    }
    async callAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr('Retrieving site administrators...');
        }
        const requestOptions = {
            url: `${args.options.siteUrl}/_api/web/siteusers?$filter=IsSiteAdmin eq true`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const responseContent = await request.get(requestOptions);
        const primaryOwnerLogin = await spo.getPrimaryOwnerLoginFromSite(args.options.siteUrl, logger, this.verbose);
        const mappedResult = responseContent.value.map((u) => ({
            Id: u.Id,
            LoginName: u.LoginName,
            Title: u.Title,
            PrincipalType: u.PrincipalType,
            PrincipalTypeString: ListPrincipalType[u.PrincipalType],
            Email: u.Email,
            IsPrimaryAdmin: u.LoginName === primaryOwnerLogin
        }));
        await logger.log(mappedResult);
    }
}
_SpoSiteAdminListCommand_instances = new WeakSet(), _SpoSiteAdminListCommand_initTelemetry = function _SpoSiteAdminListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            asAdmin: !!args.options.asAdmin
        });
    });
}, _SpoSiteAdminListCommand_initOptions = function _SpoSiteAdminListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --siteUrl <siteUrl>'
    }, {
        option: '--asAdmin'
    });
}, _SpoSiteAdminListCommand_initValidators = function _SpoSiteAdminListCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.siteUrl));
}, _SpoSiteAdminListCommand_initTypes = function _SpoSiteAdminListCommand_initTypes() {
    this.types.string.push('siteUrl');
    this.types.boolean.push('asAdmin');
};
export default new SpoSiteAdminListCommand();
//# sourceMappingURL=site-admin-list.js.map