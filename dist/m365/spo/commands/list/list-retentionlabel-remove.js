var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListRetentionLabelRemoveCommand_instances, _SpoListRetentionLabelRemoveCommand_initTelemetry, _SpoListRetentionLabelRemoveCommand_initOptions, _SpoListRetentionLabelRemoveCommand_initValidators, _SpoListRetentionLabelRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { spo } from '../../../../utils/spo.js';
class SpoListRetentionLabelRemoveCommand extends SpoCommand {
    get name() {
        return commands.LIST_RETENTIONLABEL_REMOVE;
    }
    get description() {
        return 'Clears the retention label on the specified list or library.';
    }
    constructor() {
        super();
        _SpoListRetentionLabelRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListRetentionLabelRemoveCommand_instances, "m", _SpoListRetentionLabelRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListRetentionLabelRemoveCommand_instances, "m", _SpoListRetentionLabelRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListRetentionLabelRemoveCommand_instances, "m", _SpoListRetentionLabelRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListRetentionLabelRemoveCommand_instances, "m", _SpoListRetentionLabelRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeListRetentionLabel(logger, args);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the retention label from list '${args.options.listId || args.options.listTitle || args.options.listUrl}'?` });
            if (result) {
                await this.removeListRetentionLabel(logger, args);
            }
        }
    }
    async removeListRetentionLabel(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Clears the retention label from list ${args.options.listId || args.options.listTitle || args.options.listUrl} in site at ${args.options.webUrl}...`);
        }
        try {
            const listServerRelativeUrl = await this.getListServerRelativeUrl(args, logger);
            const listAbsoluteUrl = urlUtil.getAbsoluteUrl(args.options.webUrl, listServerRelativeUrl);
            await spo.removeDefaultRetentionLabelFromList(args.options.webUrl, listAbsoluteUrl, logger, args.options.verbose);
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
_SpoListRetentionLabelRemoveCommand_instances = new WeakSet(), _SpoListRetentionLabelRemoveCommand_initTelemetry = function _SpoListRetentionLabelRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            force: !!args.options.force
        });
    });
}, _SpoListRetentionLabelRemoveCommand_initOptions = function _SpoListRetentionLabelRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-t, --listTitle [listTitle]'
    }, {
        option: '-i, --listId [listId]'
    }, {
        option: '-l, --listUrl [listUrl]'
    }, {
        option: '-f, --force'
    });
}, _SpoListRetentionLabelRemoveCommand_initValidators = function _SpoListRetentionLabelRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} is not a valid GUID`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
}, _SpoListRetentionLabelRemoveCommand_initOptionSets = function _SpoListRetentionLabelRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListRetentionLabelRemoveCommand();
//# sourceMappingURL=list-retentionlabel-remove.js.map