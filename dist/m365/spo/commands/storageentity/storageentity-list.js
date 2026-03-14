var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoStorageEntityListCommand_instances, _SpoStorageEntityListCommand_initOptions, _SpoStorageEntityListCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoStorageEntityListCommand extends SpoCommand {
    get name() {
        return commands.STORAGEENTITY_LIST;
    }
    get description() {
        return 'Lists tenant properties stored on the specified SharePoint Online app catalog';
    }
    constructor() {
        super();
        _SpoStorageEntityListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoStorageEntityListCommand_instances, "m", _SpoStorageEntityListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoStorageEntityListCommand_instances, "m", _SpoStorageEntityListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving details for all tenant properties in ${args.options.appCatalogUrl}...`);
        }
        const requestOptions = {
            url: `${args.options.appCatalogUrl}/_api/web/AllProperties?$select=storageentitiesindex`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        try {
            const web = await request.get(requestOptions);
            if (!web.storageentitiesindex ||
                web.storageentitiesindex.trim().length === 0) {
                if (this.verbose) {
                    await logger.logToStderr('No tenant properties found');
                }
            }
            else {
                const properties = JSON.parse(web.storageentitiesindex);
                const keys = Object.keys(properties);
                if (keys.length === 0) {
                    if (this.verbose) {
                        await logger.logToStderr('No tenant properties found');
                    }
                }
                else {
                    await logger.log(keys.map((key) => {
                        const property = properties[key];
                        return {
                            Key: key,
                            Value: property.Value,
                            Description: property.Description,
                            Comment: property.Comment
                        };
                    }));
                }
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoStorageEntityListCommand_instances = new WeakSet(), _SpoStorageEntityListCommand_initOptions = function _SpoStorageEntityListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --appCatalogUrl <appCatalogUrl>'
    });
}, _SpoStorageEntityListCommand_initValidators = function _SpoStorageEntityListCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.appCatalogUrl));
};
export default new SpoStorageEntityListCommand();
//# sourceMappingURL=storageentity-list.js.map