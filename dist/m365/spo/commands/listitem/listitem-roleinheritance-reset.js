var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListItemRoleInheritanceResetCommand_instances, _SpoListItemRoleInheritanceResetCommand_initTelemetry, _SpoListItemRoleInheritanceResetCommand_initOptions, _SpoListItemRoleInheritanceResetCommand_initValidators, _SpoListItemRoleInheritanceResetCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListItemRoleInheritanceResetCommand extends SpoCommand {
    get name() {
        return commands.LISTITEM_ROLEINHERITANCE_RESET;
    }
    get description() {
        return 'Restores the role inheritance of list item, file, or folder';
    }
    constructor() {
        super();
        _SpoListItemRoleInheritanceResetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListItemRoleInheritanceResetCommand_instances, "m", _SpoListItemRoleInheritanceResetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListItemRoleInheritanceResetCommand_instances, "m", _SpoListItemRoleInheritanceResetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListItemRoleInheritanceResetCommand_instances, "m", _SpoListItemRoleInheritanceResetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListItemRoleInheritanceResetCommand_instances, "m", _SpoListItemRoleInheritanceResetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Restore role inheritance of list item in site at ${args.options.webUrl}...`);
        }
        if (args.options.force) {
            await this.resetListItemRoleInheritance(args.options);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to reset the role inheritance of ${args.options.listItemId} in list ${args.options.listId ?? args.options.listTitle}?` });
            if (result) {
                await this.resetListItemRoleInheritance(args.options);
            }
        }
    }
    async resetListItemRoleInheritance(options) {
        try {
            let requestUrl = `${options.webUrl}/_api/web`;
            if (options.listId) {
                requestUrl += `/lists(guid'${formatting.encodeQueryParameter(options.listId)}')`;
            }
            else if (options.listTitle) {
                requestUrl += `/lists/getbytitle('${formatting.encodeQueryParameter(options.listTitle)}')`;
            }
            else if (options.listUrl) {
                const listServerRelativeUrl = urlUtil.getServerRelativePath(options.webUrl, options.listUrl);
                requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
            }
            const requestOptions = {
                url: `${requestUrl}/items(${options.listItemId})/resetroleinheritance`,
                method: 'POST',
                headers: {
                    'accept': 'application/json;odata=nometadata',
                    'content-type': 'application/json'
                },
                responseType: 'json'
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoListItemRoleInheritanceResetCommand_instances = new WeakSet(), _SpoListItemRoleInheritanceResetCommand_initTelemetry = function _SpoListItemRoleInheritanceResetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            force: (!(!args.options.force)).toString()
        });
    });
}, _SpoListItemRoleInheritanceResetCommand_initOptions = function _SpoListItemRoleInheritanceResetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--listItemId <listItemId>'
    }, {
        option: '--listId [listId]'
    }, {
        option: '--listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '-f, --force'
    });
}, _SpoListItemRoleInheritanceResetCommand_initValidators = function _SpoListItemRoleInheritanceResetCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} is not a valid GUID`;
        }
        if (isNaN(args.options.listItemId)) {
            return `${args.options.listItemId} is not a number`;
        }
        return true;
    });
}, _SpoListItemRoleInheritanceResetCommand_initOptionSets = function _SpoListItemRoleInheritanceResetCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListItemRoleInheritanceResetCommand();
//# sourceMappingURL=listitem-roleinheritance-reset.js.map