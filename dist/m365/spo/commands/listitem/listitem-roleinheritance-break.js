var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListItemRoleInheritanceBreakCommand_instances, _SpoListItemRoleInheritanceBreakCommand_initTelemetry, _SpoListItemRoleInheritanceBreakCommand_initOptions, _SpoListItemRoleInheritanceBreakCommand_initValidators, _SpoListItemRoleInheritanceBreakCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListItemRoleInheritanceBreakCommand extends SpoCommand {
    get name() {
        return commands.LISTITEM_ROLEINHERITANCE_BREAK;
    }
    get description() {
        return 'Break inheritance of list item';
    }
    constructor() {
        super();
        _SpoListItemRoleInheritanceBreakCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListItemRoleInheritanceBreakCommand_instances, "m", _SpoListItemRoleInheritanceBreakCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListItemRoleInheritanceBreakCommand_instances, "m", _SpoListItemRoleInheritanceBreakCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListItemRoleInheritanceBreakCommand_instances, "m", _SpoListItemRoleInheritanceBreakCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListItemRoleInheritanceBreakCommand_instances, "m", _SpoListItemRoleInheritanceBreakCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Breaking role inheritance of list item in site at ${args.options.webUrl}...`);
        }
        if (args.options.force) {
            await this.breakListItemRoleInheritance(args.options);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to break the role inheritance of ${args.options.listItemId} in list ${args.options.listId ?? args.options.listTitle}?` });
            if (result) {
                await this.breakListItemRoleInheritance(args.options);
            }
        }
    }
    async breakListItemRoleInheritance(options) {
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
            let keepExistingPermissions = true;
            if (options.clearExistingPermissions) {
                keepExistingPermissions = !options.clearExistingPermissions;
            }
            const requestOptions = {
                url: `${requestUrl}/items(${options.listItemId})/breakroleinheritance(${keepExistingPermissions})`,
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
_SpoListItemRoleInheritanceBreakCommand_instances = new WeakSet(), _SpoListItemRoleInheritanceBreakCommand_initTelemetry = function _SpoListItemRoleInheritanceBreakCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            clearExistingPermissions: args.options.clearExistingPermissions === true,
            force: (!(!args.options.force)).toString()
        });
    });
}, _SpoListItemRoleInheritanceBreakCommand_initOptions = function _SpoListItemRoleInheritanceBreakCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--listItemId <listItemId>'
    }, {
        option: '-l, --listId [listId]'
    }, {
        option: '-t, --listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '-c, --clearExistingPermissions'
    }, {
        option: '-f, --force'
    });
}, _SpoListItemRoleInheritanceBreakCommand_initValidators = function _SpoListItemRoleInheritanceBreakCommand_initValidators() {
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
}, _SpoListItemRoleInheritanceBreakCommand_initOptionSets = function _SpoListItemRoleInheritanceBreakCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListItemRoleInheritanceBreakCommand();
//# sourceMappingURL=listitem-roleinheritance-break.js.map