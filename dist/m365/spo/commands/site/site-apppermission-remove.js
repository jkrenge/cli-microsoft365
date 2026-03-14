var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteAppPermissionRemoveCommand_instances, _SpoSiteAppPermissionRemoveCommand_initTelemetry, _SpoSiteAppPermissionRemoveCommand_initOptions, _SpoSiteAppPermissionRemoveCommand_initValidators, _SpoSiteAppPermissionRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class SpoSiteAppPermissionRemoveCommand extends GraphCommand {
    get name() {
        return commands.SITE_APPPERMISSION_REMOVE;
    }
    get description() {
        return 'Removes a specific application permission from a site';
    }
    constructor() {
        super();
        _SpoSiteAppPermissionRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteAppPermissionRemoveCommand_instances, "m", _SpoSiteAppPermissionRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteAppPermissionRemoveCommand_instances, "m", _SpoSiteAppPermissionRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteAppPermissionRemoveCommand_instances, "m", _SpoSiteAppPermissionRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoSiteAppPermissionRemoveCommand_instances, "m", _SpoSiteAppPermissionRemoveCommand_initOptionSets).call(this);
    }
    getPermissions(siteId) {
        const requestOptions = {
            url: `${this.resource}/v1.0/sites/${siteId}/permissions`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        return request.get(requestOptions);
    }
    getFilteredPermissions(options, permissions) {
        let filterProperty = 'displayName';
        let filterValue = options.appDisplayName;
        if (options.appId) {
            filterProperty = 'id';
            filterValue = options.appId;
        }
        return permissions.filter((p) => {
            return p.grantedToIdentities.some(({ application }) => application[filterProperty] === filterValue);
        });
    }
    async getPermissionIds(siteId, options) {
        if (options.id) {
            return [options.id];
        }
        const permissionsObject = await this.getPermissions(siteId);
        let permissions = permissionsObject.value;
        if (options.appId || options.appDisplayName) {
            permissions = this.getFilteredPermissions(options, permissionsObject.value);
        }
        return permissions.map(x => x.id);
    }
    removePermissions(siteId, permissionId) {
        const spRequestOptions = {
            url: `${this.resource}/v1.0/sites/${siteId}/permissions/${permissionId}`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        return request.delete(spRequestOptions);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeSiteAppPermission(args.options);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the specified application permission from site ${args.options.siteUrl}?` });
            if (result) {
                await this.removeSiteAppPermission(args.options);
            }
        }
    }
    async removeSiteAppPermission(options) {
        try {
            const siteId = await spo.getSpoGraphSiteId(options.siteUrl);
            const permissionIdsToRemove = await this.getPermissionIds(siteId, options);
            const tasks = [];
            for (const permissionId of permissionIdsToRemove) {
                tasks.push(this.removePermissions(siteId, permissionId));
            }
            await Promise.all(tasks);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoSiteAppPermissionRemoveCommand_instances = new WeakSet(), _SpoSiteAppPermissionRemoveCommand_initTelemetry = function _SpoSiteAppPermissionRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            appId: typeof args.options.appId !== 'undefined',
            appDisplayName: typeof args.options.appDisplayName !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            force: !!args.options.force
        });
    });
}, _SpoSiteAppPermissionRemoveCommand_initOptions = function _SpoSiteAppPermissionRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --siteUrl <siteUrl>'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '--appId [appId]'
    }, {
        option: '-n, --appDisplayName [appDisplayName]'
    }, {
        option: '-f, --force'
    });
}, _SpoSiteAppPermissionRemoveCommand_initValidators = function _SpoSiteAppPermissionRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.appId && !validation.isValidGuid(args.options.appId)) {
            return `${args.options.appId} is not a valid GUID`;
        }
        return validation.isValidSharePointUrl(args.options.siteUrl);
    });
}, _SpoSiteAppPermissionRemoveCommand_initOptionSets = function _SpoSiteAppPermissionRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['appId', 'appDisplayName', 'id'] });
};
export default new SpoSiteAppPermissionRemoveCommand();
//# sourceMappingURL=site-apppermission-remove.js.map