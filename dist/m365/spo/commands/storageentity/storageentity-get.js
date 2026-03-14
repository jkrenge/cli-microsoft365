var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoStorageEntityGetCommand_instances, _SpoStorageEntityGetCommand_initOptions;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoStorageEntityGetCommand extends SpoCommand {
    get name() {
        return commands.STORAGEENTITY_GET;
    }
    get description() {
        return 'Get details for the specified tenant property';
    }
    constructor() {
        super();
        _SpoStorageEntityGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoStorageEntityGetCommand_instances, "m", _SpoStorageEntityGetCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoUrl = await spo.getSpoUrl(logger, this.debug);
            const requestOptions = {
                url: `${spoUrl}/_api/web/GetStorageEntity('${formatting.encodeQueryParameter(args.options.key)}')`,
                headers: {
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            const property = await request.get(requestOptions);
            if (property["odata.null"] === true) {
                if (this.verbose) {
                    await logger.logToStderr(`Property with key ${args.options.key} not found`);
                }
            }
            else {
                await logger.log({
                    Key: args.options.key,
                    Value: property.Value,
                    Description: property.Description,
                    Comment: property.Comment
                });
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoStorageEntityGetCommand_instances = new WeakSet(), _SpoStorageEntityGetCommand_initOptions = function _SpoStorageEntityGetCommand_initOptions() {
    this.options.unshift({
        option: '-k, --key <key>'
    });
};
export default new SpoStorageEntityGetCommand();
//# sourceMappingURL=storageentity-get.js.map