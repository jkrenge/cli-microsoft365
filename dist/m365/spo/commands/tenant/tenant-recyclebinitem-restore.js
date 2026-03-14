var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoTenantRecycleBinItemRestoreCommand_instances, _SpoTenantRecycleBinItemRestoreCommand_initOptions, _SpoTenantRecycleBinItemRestoreCommand_initValidators, _SpoTenantRecycleBinItemRestoreCommand_initTypes;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoTenantRecycleBinItemRestoreCommand extends SpoCommand {
    get name() {
        return commands.TENANT_RECYCLEBINITEM_RESTORE;
    }
    get description() {
        return 'Restores the specified deleted site collection from tenant recycle bin';
    }
    constructor() {
        super();
        _SpoTenantRecycleBinItemRestoreCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoTenantRecycleBinItemRestoreCommand_instances, "m", _SpoTenantRecycleBinItemRestoreCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoTenantRecycleBinItemRestoreCommand_instances, "m", _SpoTenantRecycleBinItemRestoreCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoTenantRecycleBinItemRestoreCommand_instances, "m", _SpoTenantRecycleBinItemRestoreCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Restoring site collection '${args.options.siteUrl}' from recycle bin.`);
            }
            const siteUrl = urlUtil.removeTrailingSlashes(args.options.siteUrl);
            const adminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            const requestOptions = {
                url: `${adminUrl}/_api/SPO.Tenant/RestoreDeletedSite`,
                headers: {
                    accept: 'application/json;odata=nometadata',
                    'content-type': 'application/json;charset=utf-8'
                },
                data: { siteUrl },
                responseType: 'json'
            };
            await request.post(requestOptions);
            const groupId = await this.getSiteGroupId(adminUrl, siteUrl);
            if (groupId && groupId !== '00000000-0000-0000-0000-000000000000') {
                if (this.verbose) {
                    await logger.logToStderr(`Restoring Microsoft 365 group with ID '${groupId}' from recycle bin.`);
                }
                const restoreOptions = {
                    url: `https://graph.microsoft.com/v1.0/directory/deletedItems/${groupId}/restore`,
                    headers: {
                        accept: 'application/json;odata.metadata=none',
                        'content-type': 'application/json'
                    },
                    responseType: 'json'
                };
                await request.post(restoreOptions);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getSiteGroupId(adminUrl, url) {
        const sites = await odata.getAllItems(`${adminUrl}/_api/web/lists/GetByTitle('DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS')/items?$filter=SiteUrl eq '${formatting.encodeQueryParameter(url)}'&$select=GroupId`);
        return sites[0].GroupId;
    }
}
_SpoTenantRecycleBinItemRestoreCommand_instances = new WeakSet(), _SpoTenantRecycleBinItemRestoreCommand_initOptions = function _SpoTenantRecycleBinItemRestoreCommand_initOptions() {
    this.options.unshift({
        option: '-u, --siteUrl <siteUrl>'
    });
}, _SpoTenantRecycleBinItemRestoreCommand_initValidators = function _SpoTenantRecycleBinItemRestoreCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.siteUrl));
}, _SpoTenantRecycleBinItemRestoreCommand_initTypes = function _SpoTenantRecycleBinItemRestoreCommand_initTypes() {
    this.types.string.push('siteUrl');
};
export default new SpoTenantRecycleBinItemRestoreCommand();
//# sourceMappingURL=tenant-recyclebinitem-restore.js.map