var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoAppPageSetCommand_instances, _SpoAppPageSetCommand_initOptions, _SpoAppPageSetCommand_initValidators;
import request from '../../../../request.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoAppPageSetCommand extends SpoCommand {
    get name() {
        return commands.APPPAGE_SET;
    }
    get description() {
        return 'Updates the single-part app page';
    }
    constructor() {
        super();
        _SpoAppPageSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoAppPageSetCommand_instances, "m", _SpoAppPageSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoAppPageSetCommand_instances, "m", _SpoAppPageSetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const requestOptions = {
            url: `${args.options.webUrl}/_api/sitepages/Pages/UpdateFullPageApp`,
            headers: {
                'content-type': 'application/json;odata=nometadata',
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json',
            data: {
                serverRelativeUrl: `${urlUtil.getServerRelativePath(args.options.webUrl, 'SitePages')}/${args.options.name}`,
                webPartDataAsJson: args.options.webPartData
            }
        };
        try {
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoAppPageSetCommand_instances = new WeakSet(), _SpoAppPageSetCommand_initOptions = function _SpoAppPageSetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-n, --name <name>'
    }, {
        option: '-d, --webPartData <webPartData>'
    });
}, _SpoAppPageSetCommand_initValidators = function _SpoAppPageSetCommand_initValidators() {
    this.validators.push(async (args) => {
        try {
            JSON.parse(args.options.webPartData);
        }
        catch (e) {
            return `Specified webPartData is not a valid JSON string. Error: ${e}`;
        }
        return true;
    });
};
export default new SpoAppPageSetCommand();
//# sourceMappingURL=apppage-set.js.map