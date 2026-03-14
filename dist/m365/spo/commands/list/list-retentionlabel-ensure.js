var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListRetentionLabelEnsureCommand_instances, _SpoListRetentionLabelEnsureCommand_initTelemetry, _SpoListRetentionLabelEnsureCommand_initOptions, _SpoListRetentionLabelEnsureCommand_initValidators, _SpoListRetentionLabelEnsureCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListRetentionLabelEnsureCommand extends SpoCommand {
    get name() {
        return commands.LIST_RETENTIONLABEL_ENSURE;
    }
    get description() {
        return 'Sets a default retention label on the specified list or library.';
    }
    constructor() {
        super();
        _SpoListRetentionLabelEnsureCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListRetentionLabelEnsureCommand_instances, "m", _SpoListRetentionLabelEnsureCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListRetentionLabelEnsureCommand_instances, "m", _SpoListRetentionLabelEnsureCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListRetentionLabelEnsureCommand_instances, "m", _SpoListRetentionLabelEnsureCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListRetentionLabelEnsureCommand_instances, "m", _SpoListRetentionLabelEnsureCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const listServerRelativeUrl = await this.getListServerRelativeUrl(args, logger);
            const listAbsoluteUrl = urlUtil.getAbsoluteUrl(args.options.webUrl, listServerRelativeUrl);
            await spo.applyDefaultRetentionLabelToList(args.options.webUrl, args.options.name, listAbsoluteUrl, args.options.syncToItems, logger, args.options.verbose);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getListServerRelativeUrl(args, logger) {
        if (this.verbose) {
            await logger.logToStderr('Getting the list server relative URL');
        }
        if (args.options.listUrl) {
            return urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
        }
        let requestUrl = `${args.options.webUrl}/_api/web/`;
        if (args.options.listId) {
            requestUrl += `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/`;
        }
        else {
            requestUrl += `lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')/`;
        }
        const requestOptions = {
            url: `${requestUrl}?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const listInstance = await request.get(requestOptions);
        return listInstance.RootFolder.ServerRelativeUrl;
    }
}
_SpoListRetentionLabelEnsureCommand_instances = new WeakSet(), _SpoListRetentionLabelEnsureCommand_initTelemetry = function _SpoListRetentionLabelEnsureCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: (!(!args.options.listId)).toString(),
            listTitle: (!(!args.options.listTitle)).toString(),
            listUrl: (!(!args.options.listUrl)).toString(),
            syncToItems: args.options.syncToItems || false
        });
    });
}, _SpoListRetentionLabelEnsureCommand_initOptions = function _SpoListRetentionLabelEnsureCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--name <name>'
    }, {
        option: '-t, --listTitle [listTitle]'
    }, {
        option: '-l, --listId [listId]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '--syncToItems'
    });
}, _SpoListRetentionLabelEnsureCommand_initValidators = function _SpoListRetentionLabelEnsureCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} is not a valid GUID`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
}, _SpoListRetentionLabelEnsureCommand_initOptionSets = function _SpoListRetentionLabelEnsureCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListRetentionLabelEnsureCommand();
//# sourceMappingURL=list-retentionlabel-ensure.js.map