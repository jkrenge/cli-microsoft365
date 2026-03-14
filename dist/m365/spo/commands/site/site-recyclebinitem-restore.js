var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteRecycleBinItemRestoreCommand_instances, _SpoSiteRecycleBinItemRestoreCommand_initOptions, _SpoSiteRecycleBinItemRestoreCommand_initValidators, _SpoSiteRecycleBinItemRestoreCommand_initTypes;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteRecycleBinItemRestoreCommand extends SpoCommand {
    get name() {
        return commands.SITE_RECYCLEBINITEM_RESTORE;
    }
    get description() {
        return 'Restores given items from the site recycle bin';
    }
    constructor() {
        super();
        _SpoSiteRecycleBinItemRestoreCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteRecycleBinItemRestoreCommand_instances, "m", _SpoSiteRecycleBinItemRestoreCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteRecycleBinItemRestoreCommand_instances, "m", _SpoSiteRecycleBinItemRestoreCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoSiteRecycleBinItemRestoreCommand_instances, "m", _SpoSiteRecycleBinItemRestoreCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Restoring items from recycle bin at ${args.options.siteUrl}...`);
        }
        const baseUrl = `${args.options.siteUrl}/_api`;
        try {
            if (args.options.ids) {
                const requestUrl = baseUrl + '/site/RecycleBin/RestoreByIds';
                const ids = formatting.splitAndTrim(args.options.ids);
                const idsChunks = [];
                while (ids.length) {
                    idsChunks.push(ids.splice(0, 20));
                }
                await Promise.all(idsChunks.map((idsChunk) => {
                    const requestOptions = {
                        url: requestUrl,
                        headers: {
                            'accept': 'application/json;odata=nometadata',
                            'content-type': 'application/json'
                        },
                        responseType: 'json',
                        data: {
                            ids: idsChunk
                        }
                    };
                    return request.post(requestOptions);
                }));
            }
            else {
                if (args.options.allPrimary && args.options.allSecondary) {
                    await this.restoreRecycleBinStage(baseUrl + '/site/RecycleBin/RestoreAll');
                }
                else if (args.options.allPrimary) {
                    await this.restoreRecycleBinStage(baseUrl + '/web/RecycleBin/RestoreAll');
                }
                else if (args.options.allSecondary) {
                    await this.restoreRecycleBinStage(baseUrl + '/site/GetRecycleBinItems(rowLimit=2000000000,itemState=2)/RestoreAll');
                }
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async restoreRecycleBinStage(requestUrl) {
        const requestOptions = {
            url: requestUrl,
            headers: {
                accept: 'application/json;odata=nometadata',
                'content-type': 'application/json'
            },
            responseType: 'json'
        };
        return request.post(requestOptions);
    }
}
_SpoSiteRecycleBinItemRestoreCommand_instances = new WeakSet(), _SpoSiteRecycleBinItemRestoreCommand_initOptions = function _SpoSiteRecycleBinItemRestoreCommand_initOptions() {
    this.options.unshift({
        option: '-u, --siteUrl <siteUrl>'
    }, {
        option: '-i, --ids [ids]'
    }, {
        option: '--allPrimary'
    }, {
        option: '--allSecondary'
    });
}, _SpoSiteRecycleBinItemRestoreCommand_initValidators = function _SpoSiteRecycleBinItemRestoreCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.siteUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.ids) {
            const invalidIds = formatting
                .splitAndTrim(args.options.ids)
                .filter(id => !validation.isValidGuid(id));
            if (invalidIds.length > 0) {
                return `The following IDs are invalid: ${invalidIds.join(', ')}`;
            }
        }
        if ((!args.options.ids && !args.options.allPrimary && !args.options.allSecondary)
            || (args.options.ids && (args.options.allPrimary || args.options.allSecondary))) {
            return `Option 'ids' cannot be used with 'allPrimary' or 'allSecondary'.`;
        }
        return true;
    });
}, _SpoSiteRecycleBinItemRestoreCommand_initTypes = function _SpoSiteRecycleBinItemRestoreCommand_initTypes() {
    this.types.string.push('siteUrl', 'ids');
    this.types.boolean.push('allPrimary', 'allSecondary');
};
export default new SpoSiteRecycleBinItemRestoreCommand();
//# sourceMappingURL=site-recyclebinitem-restore.js.map