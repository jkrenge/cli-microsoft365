var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteAppPermissionListCommand_instances, _SpoSiteAppPermissionListCommand_initTelemetry, _SpoSiteAppPermissionListCommand_initOptions, _SpoSiteAppPermissionListCommand_initValidators;
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class SpoSiteAppPermissionListCommand extends GraphCommand {
    get name() {
        return commands.SITE_APPPERMISSION_LIST;
    }
    get description() {
        return 'Lists application permissions for a site';
    }
    constructor() {
        super();
        _SpoSiteAppPermissionListCommand_instances.add(this);
        this.siteId = '';
        __classPrivateFieldGet(this, _SpoSiteAppPermissionListCommand_instances, "m", _SpoSiteAppPermissionListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteAppPermissionListCommand_instances, "m", _SpoSiteAppPermissionListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteAppPermissionListCommand_instances, "m", _SpoSiteAppPermissionListCommand_initValidators).call(this);
    }
    getFilteredPermissions(args, permissions) {
        let filterProperty = 'displayName';
        let filterValue = args.options.appDisplayName;
        if (args.options.appId) {
            filterProperty = 'id';
            filterValue = args.options.appId;
        }
        return permissions.filter((p) => {
            return p.grantedToIdentities.some(({ application }) => application[filterProperty] === filterValue);
        });
    }
    getApplicationPermission(permissionId) {
        const requestOptions = {
            url: `${this.resource}/v1.0/sites/${this.siteId}/permissions/${permissionId}`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        return request.get(requestOptions);
    }
    getTransposed(permissions) {
        const transposed = [];
        permissions.forEach((permissionObject) => {
            permissionObject.grantedToIdentities.forEach((permissionEntity) => {
                transposed.push({
                    appDisplayName: permissionEntity.application.displayName,
                    appId: permissionEntity.application.id,
                    permissionId: permissionObject.id,
                    roles: permissionObject.roles
                });
            });
        });
        return transposed;
    }
    getPermissions() {
        const requestOptions = {
            url: `${this.resource}/v1.0/sites/${this.siteId}/permissions`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        return request.get(requestOptions);
    }
    async commandAction(logger, args) {
        try {
            this.siteId = await spo.getSpoGraphSiteId(args.options.siteUrl);
            const permRes = await this.getPermissions();
            let permissions = permRes.value;
            if (args.options.appId || args.options.appDisplayName) {
                permissions = this.getFilteredPermissions(args, permRes.value);
            }
            const res = await Promise.all(permissions.map(g => this.getApplicationPermission(g.id)));
            await logger.log(this.getTransposed(res));
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoSiteAppPermissionListCommand_instances = new WeakSet(), _SpoSiteAppPermissionListCommand_initTelemetry = function _SpoSiteAppPermissionListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            appDisplayName: typeof args.options.appDisplayName !== 'undefined',
            appId: typeof args.options.appId !== 'undefined'
        });
    });
}, _SpoSiteAppPermissionListCommand_initOptions = function _SpoSiteAppPermissionListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --siteUrl <siteUrl>'
    }, {
        option: '-i, --appId [appId]'
    }, {
        option: '-n, --appDisplayName [appDisplayName]'
    });
}, _SpoSiteAppPermissionListCommand_initValidators = function _SpoSiteAppPermissionListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.appId && args.options.appDisplayName) {
            return `Provide either appId or appDisplayName, not both`;
        }
        return validation.isValidSharePointUrl(args.options.siteUrl);
    });
};
export default new SpoSiteAppPermissionListCommand();
//# sourceMappingURL=site-apppermission-list.js.map