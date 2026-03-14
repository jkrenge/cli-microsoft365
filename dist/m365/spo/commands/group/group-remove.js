var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoGroupRemoveCommand_instances, _SpoGroupRemoveCommand_initTelemetry, _SpoGroupRemoveCommand_initOptions, _SpoGroupRemoveCommand_initValidators, _SpoGroupRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoGroupRemoveCommand extends SpoCommand {
    get name() {
        return commands.GROUP_REMOVE;
    }
    get description() {
        return 'Removes group from specific web';
    }
    constructor() {
        super();
        _SpoGroupRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoGroupRemoveCommand_instances, "m", _SpoGroupRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoGroupRemoveCommand_instances, "m", _SpoGroupRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoGroupRemoveCommand_instances, "m", _SpoGroupRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoGroupRemoveCommand_instances, "m", _SpoGroupRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const removeGroup = async () => {
            if (this.verbose) {
                await logger.logToStderr(`Removing group in web at ${args.options.webUrl}...`);
            }
            try {
                let groupId;
                if (args.options.name) {
                    const requestOptions = {
                        url: `${args.options.webUrl}/_api/web/sitegroups/GetByName('${args.options.name}')?$select=Id`,
                        headers: {
                            accept: 'application/json'
                        },
                        responseType: 'json'
                    };
                    const group = await request.get(requestOptions);
                    groupId = group.Id;
                }
                else {
                    groupId = args.options.id;
                }
                const requestUrl = `${args.options.webUrl}/_api/web/sitegroups/RemoveById(${groupId})`;
                const requestOptions = {
                    url: requestUrl,
                    method: 'POST',
                    headers: {
                        'content-length': 0,
                        'accept': 'application/json'
                    },
                    responseType: 'json'
                };
                await request.post(requestOptions);
                // REST post call doesn't return anything
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeGroup();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the group ${args.options.id || args.options.name} from web ${args.options.webUrl}?` });
            if (result) {
                await removeGroup();
            }
        }
    }
}
_SpoGroupRemoveCommand_instances = new WeakSet(), _SpoGroupRemoveCommand_initTelemetry = function _SpoGroupRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: (!(!args.options.id)).toString(),
            name: (!(!args.options.name)).toString(),
            force: (!(!args.options.force)).toString()
        });
    });
}, _SpoGroupRemoveCommand_initOptions = function _SpoGroupRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--id [id]'
    }, {
        option: '--name [name]'
    }, {
        option: '-f, --force'
    });
}, _SpoGroupRemoveCommand_initValidators = function _SpoGroupRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.id && typeof args.options.id !== 'number') {
            return `${args.options.id} is not a number`;
        }
        return true;
    });
}, _SpoGroupRemoveCommand_initOptionSets = function _SpoGroupRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] });
};
export default new SpoGroupRemoveCommand();
//# sourceMappingURL=group-remove.js.map