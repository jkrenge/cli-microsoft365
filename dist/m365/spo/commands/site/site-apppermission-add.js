var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteAppPermissionAddCommand_instances, _SpoSiteAppPermissionAddCommand_initTelemetry, _SpoSiteAppPermissionAddCommand_initOptions, _SpoSiteAppPermissionAddCommand_initValidators, _SpoSiteAppPermissionAddCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class SpoSiteAppPermissionAddCommand extends GraphCommand {
    get name() {
        return commands.SITE_APPPERMISSION_ADD;
    }
    get description() {
        return 'Adds an application permissions to the site';
    }
    constructor() {
        super();
        _SpoSiteAppPermissionAddCommand_instances.add(this);
        this.siteId = '';
        this.roles = ['read', 'write', 'manage', 'fullcontrol'];
        __classPrivateFieldGet(this, _SpoSiteAppPermissionAddCommand_instances, "m", _SpoSiteAppPermissionAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteAppPermissionAddCommand_instances, "m", _SpoSiteAppPermissionAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteAppPermissionAddCommand_instances, "m", _SpoSiteAppPermissionAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoSiteAppPermissionAddCommand_instances, "m", _SpoSiteAppPermissionAddCommand_initOptionSets).call(this);
    }
    async getAppInfo(args) {
        if (args.options.appId && args.options.appDisplayName) {
            return {
                appId: args.options.appId,
                displayName: args.options.appDisplayName
            };
        }
        let endpoint;
        if (args.options.appId) {
            endpoint = `${this.resource}/v1.0/myorganization/servicePrincipals?$select=appId,displayName&$filter=appId eq '${formatting.encodeQueryParameter(args.options.appId)}'`;
        }
        else {
            endpoint = `${this.resource}/v1.0/myorganization/servicePrincipals?$select=appId,displayName&$filter=displayName eq '${formatting.encodeQueryParameter(args.options.appDisplayName)}'`;
        }
        const appRequestOptions = {
            url: endpoint,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const response = await request.get(appRequestOptions);
        const appItem = response.value[0];
        if (!appItem) {
            throw "The specified Microsoft Entra app does not exist";
        }
        if (response.value.length > 1) {
            throw `Multiple Microsoft Entra apps with displayName ${args.options.appDisplayName} found: ${response.value.map(x => x.appId)}`;
        }
        return {
            appId: appItem.appId,
            displayName: appItem.displayName
        };
    }
    /**
     * Checks if the requested permission needs elevation after the initial creation.
     */
    roleNeedsElevation(permission) {
        return ['manage', 'fullcontrol'].indexOf(permission) > -1;
    }
    /**
     * Grants the app 'read' or 'write' permissions to the site.
     *
     * Explanation:
     * 'manage' and 'fullcontrol' permissions cannot be granted directly when adding app permissions.
     * They can currently only be assigned when updating existing app permissions.
     * We therefore assign 'write' permissions first, and update it to the requested role afterwards.
     */
    addPermissions(args, appInfo) {
        const requestOptions = {
            url: `${this.resource}/v1.0/sites/${this.siteId}/permissions`,
            headers: {
                accept: 'application/json;odata.metadata=none',
                'content-type': 'application/json;odata=nometadata'
            },
            data: {
                roles: [this.roleNeedsElevation(args.options.permission) ? 'write' : args.options.permission],
                grantedToIdentities: [{ application: { "id": appInfo.appId, "displayName": appInfo.displayName } }]
            },
            responseType: 'json'
        };
        return request.post(requestOptions);
    }
    /**
     * Updates the granted permissions to 'manage' or 'fullcontrol'.
     */
    elevatePermissions(args, permission) {
        const requestOptions = {
            url: `${this.resource}/v1.0/sites/${this.siteId}/permissions/${permission.id}`,
            headers: {
                accept: 'application/json;odata.metadata=none',
                'content-type': 'application/json;odata=nometadata'
            },
            data: {
                roles: [args.options.permission]
            },
            responseType: 'json'
        };
        return request.patch(requestOptions);
    }
    async commandAction(logger, args) {
        try {
            this.siteId = await spo.getSpoGraphSiteId(args.options.siteUrl);
            const appInfo = await this.getAppInfo(args);
            let permission = await this.addPermissions(args, appInfo);
            if (this.roleNeedsElevation(args.options.permission)) {
                permission = await this.elevatePermissions(args, permission);
            }
            await logger.log(permission);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoSiteAppPermissionAddCommand_instances = new WeakSet(), _SpoSiteAppPermissionAddCommand_initTelemetry = function _SpoSiteAppPermissionAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            appId: typeof args.options.appId !== 'undefined',
            appDisplayName: typeof args.options.appDisplayName !== 'undefined',
            permissions: args.options.permissions
        });
    });
}, _SpoSiteAppPermissionAddCommand_initOptions = function _SpoSiteAppPermissionAddCommand_initOptions() {
    this.options.unshift({
        option: '-u, --siteUrl <siteUrl>'
    }, {
        option: '-p, --permission <permission>',
        autocomplete: this.roles
    }, {
        option: '-i, --appId [appId]'
    }, {
        option: '-n, --appDisplayName [appDisplayName]'
    });
}, _SpoSiteAppPermissionAddCommand_initValidators = function _SpoSiteAppPermissionAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.appId && !validation.isValidGuid(args.options.appId)) {
            return `${args.options.appId} is not a valid GUID`;
        }
        if (this.roles.indexOf(args.options.permission) === -1) {
            return `${args.options.permission} is not a valid permission value. Allowed values are ${this.roles.join('|')}`;
        }
        return validation.isValidSharePointUrl(args.options.siteUrl);
    });
}, _SpoSiteAppPermissionAddCommand_initOptionSets = function _SpoSiteAppPermissionAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['appId', 'appDisplayName'], runsWhen: (args) => !args.options.appId && !args.options.appDisplayName });
};
export default new SpoSiteAppPermissionAddCommand();
//# sourceMappingURL=site-apppermission-add.js.map