var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFieldListCommand_instances, _SpoFieldListCommand_initTelemetry, _SpoFieldListCommand_initOptions, _SpoFieldListCommand_initValidators;
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFieldListCommand extends SpoCommand {
    get name() {
        return commands.FIELD_LIST;
    }
    get description() {
        return 'Retrieves columns for the specified list or site';
    }
    defaultProperties() {
        return ['Id', 'Title', 'InternalName', 'Hidden'];
    }
    constructor() {
        super();
        _SpoFieldListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFieldListCommand_instances, "m", _SpoFieldListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFieldListCommand_instances, "m", _SpoFieldListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFieldListCommand_instances, "m", _SpoFieldListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        let listUrl = '';
        if (args.options.listId) {
            listUrl = `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/`;
        }
        else if (args.options.listTitle) {
            listUrl = `lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')/`;
        }
        else if (args.options.listUrl) {
            const listServerRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
            listUrl = `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/`;
        }
        try {
            const res = await odata.getAllItems(`${args.options.webUrl}/_api/web/${listUrl}fields`);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoFieldListCommand_instances = new WeakSet(), _SpoFieldListCommand_initTelemetry = function _SpoFieldListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined'
        });
    });
}, _SpoFieldListCommand_initOptions = function _SpoFieldListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-t, --listTitle [listTitle]'
    }, {
        option: '-i, --listId [listId]'
    }, {
        option: '--listUrl [listUrl]'
    });
}, _SpoFieldListCommand_initValidators = function _SpoFieldListCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} is not a valid GUID`;
        }
        const listOptions = [args.options.listId, args.options.listTitle, args.options.listUrl];
        if (listOptions.some(item => item !== undefined) && listOptions.filter(item => item !== undefined).length > 1) {
            return `Specify either list id or title or list url`;
        }
        return true;
    });
};
export default new SpoFieldListCommand();
//# sourceMappingURL=field-list.js.map