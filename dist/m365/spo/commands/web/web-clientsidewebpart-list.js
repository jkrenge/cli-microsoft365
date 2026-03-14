var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoWebClientSideWebPartListCommand_instances, _SpoWebClientSideWebPartListCommand_initOptions, _SpoWebClientSideWebPartListCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoWebClientSideWebPartListCommand extends SpoCommand {
    get name() {
        return commands.WEB_CLIENTSIDEWEBPART_LIST;
    }
    get description() {
        return 'Lists available client-side web parts';
    }
    constructor() {
        super();
        _SpoWebClientSideWebPartListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoWebClientSideWebPartListCommand_instances, "m", _SpoWebClientSideWebPartListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoWebClientSideWebPartListCommand_instances, "m", _SpoWebClientSideWebPartListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const requestOptions = {
            url: `${args.options.webUrl}/_api/web/GetClientSideWebParts`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        try {
            const res = await request.get(requestOptions);
            const clientSideWebParts = [];
            res.value.forEach(component => {
                if (component.ComponentType === 1) {
                    clientSideWebParts.push({
                        Id: component.Id.replace("{", "").replace("}", ""),
                        Name: component.Name,
                        Title: JSON.parse(component.Manifest).preconfiguredEntries[0].title.default
                    });
                }
            });
            await logger.log(clientSideWebParts);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoWebClientSideWebPartListCommand_instances = new WeakSet(), _SpoWebClientSideWebPartListCommand_initOptions = function _SpoWebClientSideWebPartListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    });
}, _SpoWebClientSideWebPartListCommand_initValidators = function _SpoWebClientSideWebPartListCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.webUrl));
};
export default new SpoWebClientSideWebPartListCommand();
//# sourceMappingURL=web-clientsidewebpart-list.js.map