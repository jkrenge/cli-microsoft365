var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFieldGetCommand_instances, _SpoFieldGetCommand_initTelemetry, _SpoFieldGetCommand_initOptions, _SpoFieldGetCommand_initValidators, _SpoFieldGetCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFieldGetCommand extends SpoCommand {
    get name() {
        return commands.FIELD_GET;
    }
    get description() {
        return 'Retrieves information about the specified list- or site column';
    }
    constructor() {
        super();
        _SpoFieldGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFieldGetCommand_instances, "m", _SpoFieldGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFieldGetCommand_instances, "m", _SpoFieldGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFieldGetCommand_instances, "m", _SpoFieldGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFieldGetCommand_instances, "m", _SpoFieldGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        let listRestUrl = '';
        if (args.options.listId) {
            listRestUrl = `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/`;
        }
        else if (args.options.listTitle) {
            listRestUrl = `lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')/`;
        }
        else if (args.options.listUrl) {
            const listServerRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
            listRestUrl = `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/`;
        }
        let fieldRestUrl;
        if (args.options.id) {
            fieldRestUrl = `/getbyid('${formatting.encodeQueryParameter(args.options.id)}')`;
        }
        else {
            fieldRestUrl = `/getbyinternalnameortitle('${formatting.encodeQueryParameter((args.options.title || args.options.internalName))}')`;
        }
        const requestOptions = {
            url: `${args.options.webUrl}/_api/web/${listRestUrl}fields${fieldRestUrl}`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        try {
            const res = await request.get(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoFieldGetCommand_instances = new WeakSet(), _SpoFieldGetCommand_initTelemetry = function _SpoFieldGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            title: typeof args.options.title !== 'undefined',
            internalName: typeof args.options.internalName !== 'undefined'
        });
    });
}, _SpoFieldGetCommand_initOptions = function _SpoFieldGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-l, --listTitle [listTitle]'
    }, {
        option: '--listId [listId]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-t, --title [title]'
    }, {
        option: '--internalName [internalName]'
    });
}, _SpoFieldGetCommand_initValidators = function _SpoFieldGetCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} is not a valid GUID`;
        }
        return true;
    });
}, _SpoFieldGetCommand_initOptionSets = function _SpoFieldGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'title', 'internalName'] });
};
export default new SpoFieldGetCommand();
//# sourceMappingURL=field-get.js.map