var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraGroupGetCommand_instances, _EntraGroupGetCommand_initOptions, _EntraGroupGetCommand_initValidators, _EntraGroupGetCommand_initOptionSets, _EntraGroupGetCommand_initTelemetry;
import { entraGroup } from '../../../../utils/entraGroup.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraGroupGetCommand extends GraphCommand {
    get name() {
        return commands.GROUP_GET;
    }
    get description() {
        return 'Gets information about the specified Entra group';
    }
    constructor() {
        super();
        _EntraGroupGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraGroupGetCommand_instances, "m", _EntraGroupGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraGroupGetCommand_instances, "m", _EntraGroupGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraGroupGetCommand_instances, "m", _EntraGroupGetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _EntraGroupGetCommand_instances, "m", _EntraGroupGetCommand_initTelemetry).call(this);
    }
    async commandAction(logger, args) {
        let group;
        try {
            if (args.options.id) {
                group = await entraGroup.getGroupById(args.options.id, args.options.properties);
            }
            else {
                group = await entraGroup.getGroupByDisplayName(args.options.displayName, args.options.properties);
            }
            await logger.log(group);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_EntraGroupGetCommand_instances = new WeakSet(), _EntraGroupGetCommand_initOptions = function _EntraGroupGetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-n, --displayName [displayName]'
    }, {
        option: '-p, --properties [properties]'
    });
}, _EntraGroupGetCommand_initValidators = function _EntraGroupGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
}, _EntraGroupGetCommand_initOptionSets = function _EntraGroupGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'displayName'] });
}, _EntraGroupGetCommand_initTelemetry = function _EntraGroupGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            displayName: typeof args.options.displayName !== 'undefined',
            properties: typeof args.options.properties !== 'undefined'
        });
    });
};
export default new EntraGroupGetCommand();
//# sourceMappingURL=group-get.js.map