var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteAppPermissionSetCommand_instances, _SpoSiteAppPermissionSetCommand_initTelemetry, _SpoSiteAppPermissionSetCommand_initOptions, _SpoSiteAppPermissionSetCommand_initValidators, _SpoSiteAppPermissionSetCommand_initOptionSets;
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class SpoSiteAppPermissionSetCommand extends GraphCommand {
    get name() {
        return commands.SITE_APPPERMISSION_SET;
    }
    get description() {
        return 'Updates a specific application permission for a site';
    }
    constructor() {
        super();
        _SpoSiteAppPermissionSetCommand_instances.add(this);
        this.siteId = '';
        this.roles = ['read', 'write', 'manage', 'fullcontrol'];
        __classPrivateFieldGet(this, _SpoSiteAppPermissionSetCommand_instances, "m", _SpoSiteAppPermissionSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteAppPermissionSetCommand_instances, "m", _SpoSiteAppPermissionSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteAppPermissionSetCommand_instances, "m", _SpoSiteAppPermissionSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoSiteAppPermissionSetCommand_instances, "m", _SpoSiteAppPermissionSetCommand_initOptionSets).call(this);
    }
    getFilteredPermissions(args, permissions) {
        let filterProperty = 'displayName';
        let filterValue = args.options.appDisplayName;
        if (args.options.appId) {
            filterProperty = 'id';
            filterValue = args.options.appId;
        }
        return permissions.filter((p) => p.grantedToIdentities.some(({ application }) => application[filterProperty] === filterValue));
    }
    async getPermission(args) {
        if (args.options.id) {
            return args.options.id;
        }
        const permissionRequestOptions = {
            url: `${this.resource}/v1.0/sites/${this.siteId}/permissions`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const response = await request.get(permissionRequestOptions);
        const sitePermissionItems = this.getFilteredPermissions(args, response.value);
        if (sitePermissionItems.length === 0) {
            throw 'The specified app permission does not exist';
        }
        if (sitePermissionItems.length > 1) {
            throw `Multiple app permissions with displayName ${args.options.appDisplayName} found: ${response.value.map(x => x.grantedToIdentities.map(y => y.application.id))}`;
        }
        return sitePermissionItems[0].id;
    }
    async commandAction(logger, args) {
        try {
            this.siteId = await spo.getSpoGraphSiteId(args.options.siteUrl);
            const sitePermissionId = await this.getPermission(args);
            const requestOptions = {
                url: `${this.resource}/v1.0/sites/${this.siteId}/permissions/${sitePermissionId}`,
                headers: {
                    accept: 'application/json;odata.metadata=none',
                    'content-type': 'application/json;odata=nometadata'
                },
                data: {
                    roles: [args.options.permission]
                },
                responseType: 'json'
            };
            const res = await request.patch(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoSiteAppPermissionSetCommand_instances = new WeakSet(), _SpoSiteAppPermissionSetCommand_initTelemetry = function _SpoSiteAppPermissionSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            appId: typeof args.options.appId !== 'undefined',
            appDisplayName: typeof args.options.appDisplayName !== 'undefined',
            permissions: args.options.permissions
        });
    });
}, _SpoSiteAppPermissionSetCommand_initOptions = function _SpoSiteAppPermissionSetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --siteUrl <siteUrl>'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '--appId [appId]'
    }, {
        option: '-n, --appDisplayName [appDisplayName]'
    }, {
        option: '-p, --permission <permission>',
        autocomplete: this.roles
    });
}, _SpoSiteAppPermissionSetCommand_initValidators = function _SpoSiteAppPermissionSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.appId && !validation.isValidGuid(args.options.appId)) {
            return `${args.options.appId} is not a valid GUID`;
        }
        if (this.roles.indexOf(args.options.permission) === -1) {
            return `${args.options.permission} is not a valid permission value. Allowed values are ${this.roles.join('|')}`;
        }
        return validation.isValidSharePointUrl(args.options.siteUrl);
    });
}, _SpoSiteAppPermissionSetCommand_initOptionSets = function _SpoSiteAppPermissionSetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'appId', 'appDisplayName'] });
};
export default new SpoSiteAppPermissionSetCommand();
//# sourceMappingURL=site-apppermission-set.js.map