var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListRemoveCommand_instances, _SpoListRemoveCommand_initTelemetry, _SpoListRemoveCommand_initOptions, _SpoListRemoveCommand_initValidators, _SpoListRemoveCommand_initOptionSets, _SpoListRemoveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListRemoveCommand extends SpoCommand {
    get name() {
        return commands.LIST_REMOVE;
    }
    get description() {
        return 'Removes the specified list';
    }
    constructor() {
        super();
        _SpoListRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListRemoveCommand_instances, "m", _SpoListRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListRemoveCommand_instances, "m", _SpoListRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListRemoveCommand_instances, "m", _SpoListRemoveCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _SpoListRemoveCommand_instances, "m", _SpoListRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListRemoveCommand_instances, "m", _SpoListRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const removeList = async () => {
            if (this.verbose) {
                await logger.logToStderr(`Removing list in site at ${args.options.webUrl}...`);
            }
            let requestUrl;
            if (args.options.id) {
                requestUrl = `${args.options.webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(args.options.id)}')`;
            }
            else {
                requestUrl = `${args.options.webUrl}/_api/web/lists/GetByTitle('${formatting.encodeQueryParameter(args.options.title)}')`;
            }
            if (args.options.recycle) {
                requestUrl += `/recycle`;
            }
            const requestOptions = {
                url: requestUrl,
                method: 'POST',
                headers: {
                    'X-HTTP-Method': 'DELETE',
                    'If-Match': '*',
                    'accept': 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            try {
                await request.post(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeList();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the list ${args.options.id || args.options.title} from site ${args.options.webUrl}?` });
            if (result) {
                await removeList();
            }
        }
    }
}
_SpoListRemoveCommand_instances = new WeakSet(), _SpoListRemoveCommand_initTelemetry = function _SpoListRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            title: typeof args.options.title !== 'undefined',
            force: !!args.options.force,
            recycle: !!args.options.recycle
        });
    });
}, _SpoListRemoveCommand_initOptions = function _SpoListRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-t, --title [title]'
    }, {
        option: '--recycle'
    }, {
        option: '-f, --force'
    });
}, _SpoListRemoveCommand_initValidators = function _SpoListRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.id &&
            !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
}, _SpoListRemoveCommand_initOptionSets = function _SpoListRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'title'] });
}, _SpoListRemoveCommand_initTypes = function _SpoListRemoveCommand_initTypes() {
    this.types.string.push('id', 'title', 'webUrl');
    this.types.boolean.push('force', 'recycle');
};
export default new SpoListRemoveCommand();
//# sourceMappingURL=list-remove.js.map