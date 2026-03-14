var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListRoleInheritanceResetCommand_instances, _SpoListRoleInheritanceResetCommand_initTelemetry, _SpoListRoleInheritanceResetCommand_initOptions, _SpoListRoleInheritanceResetCommand_initValidators, _SpoListRoleInheritanceResetCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListRoleInheritanceResetCommand extends SpoCommand {
    get name() {
        return commands.LIST_ROLEINHERITANCE_RESET;
    }
    get description() {
        return 'Restores role inheritance on list or library';
    }
    constructor() {
        super();
        _SpoListRoleInheritanceResetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListRoleInheritanceResetCommand_instances, "m", _SpoListRoleInheritanceResetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListRoleInheritanceResetCommand_instances, "m", _SpoListRoleInheritanceResetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListRoleInheritanceResetCommand_instances, "m", _SpoListRoleInheritanceResetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListRoleInheritanceResetCommand_instances, "m", _SpoListRoleInheritanceResetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Restore role inheritance of list in site at ${args.options.webUrl}...`);
        }
        const resetListRoleInheritance = async () => {
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
                const requestOptions = {
                    url: `${requestUrl}resetroleinheritance`,
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
            await resetListRoleInheritance();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to reset the role inheritance of ${args.options.listId ?? args.options.listTitle}?` });
            if (result) {
                await resetListRoleInheritance();
            }
        }
    }
}
_SpoListRoleInheritanceResetCommand_instances = new WeakSet(), _SpoListRoleInheritanceResetCommand_initTelemetry = function _SpoListRoleInheritanceResetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            force: (!(!args.options.force)).toString()
        });
    });
}, _SpoListRoleInheritanceResetCommand_initOptions = function _SpoListRoleInheritanceResetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-i, --listId [listId]'
    }, {
        option: '-t, --listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '-f, --force'
    });
}, _SpoListRoleInheritanceResetCommand_initValidators = function _SpoListRoleInheritanceResetCommand_initValidators() {
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
}, _SpoListRoleInheritanceResetCommand_initOptionSets = function _SpoListRoleInheritanceResetCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListRoleInheritanceResetCommand();
//# sourceMappingURL=list-roleinheritance-reset.js.map