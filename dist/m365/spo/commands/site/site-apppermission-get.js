var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteAppPermissionGetCommand_instances, _SpoSiteAppPermissionGetCommand_initOptions, _SpoSiteAppPermissionGetCommand_initValidators;
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class SpoSiteAppPermissionGetCommand extends GraphCommand {
    get name() {
        return commands.SITE_APPPERMISSION_GET;
    }
    get description() {
        return 'Get a specific application permissions for the site';
    }
    constructor() {
        super();
        _SpoSiteAppPermissionGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteAppPermissionGetCommand_instances, "m", _SpoSiteAppPermissionGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteAppPermissionGetCommand_instances, "m", _SpoSiteAppPermissionGetCommand_initValidators).call(this);
    }
    getApplicationPermission(args, siteId) {
        const requestOptions = {
            url: `${this.resource}/v1.0/sites/${siteId}/permissions/${args.options.id}`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        return request.get(requestOptions);
    }
    async commandAction(logger, args) {
        try {
            const siteId = await spo.getSpoGraphSiteId(args.options.siteUrl);
            const permissionObject = await this.getApplicationPermission(args, siteId);
            const transposed = [];
            permissionObject.grantedToIdentities.forEach((permissionEntity) => {
                transposed.push({
                    appDisplayName: permissionEntity.application.displayName,
                    appId: permissionEntity.application.id,
                    permissionId: permissionObject.id,
                    roles: permissionObject.roles.join()
                });
            });
            await logger.log(transposed);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoSiteAppPermissionGetCommand_instances = new WeakSet(), _SpoSiteAppPermissionGetCommand_initOptions = function _SpoSiteAppPermissionGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --siteUrl <siteUrl>'
    }, {
        option: '-i, --id <id>'
    });
}, _SpoSiteAppPermissionGetCommand_initValidators = function _SpoSiteAppPermissionGetCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.siteUrl));
};
export default new SpoSiteAppPermissionGetCommand();
//# sourceMappingURL=site-apppermission-get.js.map