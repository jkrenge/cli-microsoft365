var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListRoleInheritanceBreakCommand_instances, _SpoListRoleInheritanceBreakCommand_initTelemetry, _SpoListRoleInheritanceBreakCommand_initOptions, _SpoListRoleInheritanceBreakCommand_initValidators, _SpoListRoleInheritanceBreakCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListRoleInheritanceBreakCommand extends SpoCommand {
    get name() {
        return commands.LIST_ROLEINHERITANCE_BREAK;
    }
    get description() {
        return 'Breaks role inheritance on list or library';
    }
    constructor() {
        super();
        _SpoListRoleInheritanceBreakCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListRoleInheritanceBreakCommand_instances, "m", _SpoListRoleInheritanceBreakCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListRoleInheritanceBreakCommand_instances, "m", _SpoListRoleInheritanceBreakCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListRoleInheritanceBreakCommand_instances, "m", _SpoListRoleInheritanceBreakCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListRoleInheritanceBreakCommand_instances, "m", _SpoListRoleInheritanceBreakCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Breaking role inheritance of list in site at ${args.options.webUrl}...`);
        }
        const breakListRoleInheritance = async () => {
            try {
                let requestUrl = `${args.options.webUrl}/_api/web/`;
                if (args.options.listId) {
                    requestUrl += `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/`;
                }
                else if (args.options.listTitle) {
                    requestUrl += `lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')/`;
                }
                else if (args.options.listUrl) {
                    const listServerRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
                    requestUrl += `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/`;
                }
                let keepExistingPermissions = true;
                if (args.options.clearExistingPermissions) {
                    keepExistingPermissions = !args.options.clearExistingPermissions;
                }
                const requestOptions = {
                    url: `${requestUrl}breakroleinheritance(${keepExistingPermissions})`,
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
        };
        if (args.options.force) {
            await breakListRoleInheritance();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to break the role inheritance of ${args.options.listId ?? args.options.listTitle}?` });
            if (result) {
                await breakListRoleInheritance();
            }
        }
    }
}
_SpoListRoleInheritanceBreakCommand_instances = new WeakSet(), _SpoListRoleInheritanceBreakCommand_initTelemetry = function _SpoListRoleInheritanceBreakCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            clearExistingPermissions: args.options.clearExistingPermissions === true,
            force: (!(!args.options.force)).toString()
        });
    });
}, _SpoListRoleInheritanceBreakCommand_initOptions = function _SpoListRoleInheritanceBreakCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-i, --listId [listId]'
    }, {
        option: '-t, --listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '-c, --clearExistingPermissions'
    }, {
        option: '-f, --force'
    });
}, _SpoListRoleInheritanceBreakCommand_initValidators = function _SpoListRoleInheritanceBreakCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} is not a valid GUID`;
        }
        return true;
    });
}, _SpoListRoleInheritanceBreakCommand_initOptionSets = function _SpoListRoleInheritanceBreakCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListRoleInheritanceBreakCommand();
//# sourceMappingURL=list-roleinheritance-break.js.map