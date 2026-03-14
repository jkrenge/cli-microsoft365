var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpeContainerListCommand_instances, _SpeContainerListCommand_initTelemetry, _SpeContainerListCommand_initOptions, _SpeContainerListCommand_initValidators, _SpeContainerListCommand_initOptionSets, _SpeContainerListCommand_initTypes;
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { spe } from '../../../../utils/spe.js';
class SpeContainerListCommand extends GraphCommand {
    get name() {
        return commands.CONTAINER_LIST;
    }
    get description() {
        return 'Lists all Container Types';
    }
    defaultProperties() {
        return ['id', 'displayName', 'containerTypeId', 'createdDateTime'];
    }
    constructor() {
        super();
        _SpeContainerListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpeContainerListCommand_instances, "m", _SpeContainerListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpeContainerListCommand_instances, "m", _SpeContainerListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpeContainerListCommand_instances, "m", _SpeContainerListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpeContainerListCommand_instances, "m", _SpeContainerListCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpeContainerListCommand_instances, "m", _SpeContainerListCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Retrieving list of Containers...`);
            }
            const containerTypeId = await this.getContainerTypeId(logger, args.options);
            const allContainers = await odata.getAllItems(`${this.resource}/v1.0/storage/fileStorage/containers?$filter=containerTypeId eq ${formatting.encodeQueryParameter(containerTypeId)}`);
            await logger.log(allContainers);
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
    async getContainerTypeId(logger, options) {
        if (options.containerTypeId) {
            return options.containerTypeId;
        }
        return spe.getContainerTypeIdByName(options.containerTypeName);
    }
}
_SpeContainerListCommand_instances = new WeakSet(), _SpeContainerListCommand_initTelemetry = function _SpeContainerListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            containerTypeId: typeof args.options.containerTypeId !== 'undefined',
            containerTypeName: typeof args.options.containerTypeName !== 'undefined'
        });
    });
}, _SpeContainerListCommand_initOptions = function _SpeContainerListCommand_initOptions() {
    this.options.unshift({
        option: '--containerTypeId [containerTypeId]'
    }, {
        option: '--containerTypeName [containerTypeName]'
    });
}, _SpeContainerListCommand_initValidators = function _SpeContainerListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.containerTypeId && !validation.isValidGuid(args.options.containerTypeId)) {
            return `${args.options.containerTypeId} is not a valid GUID`;
        }
        return true;
    });
}, _SpeContainerListCommand_initOptionSets = function _SpeContainerListCommand_initOptionSets() {
    this.optionSets.push({ options: ['containerTypeId', 'containerTypeName'] });
}, _SpeContainerListCommand_initTypes = function _SpeContainerListCommand_initTypes() {
    this.types.string.push('containerTypeId', 'containerTypeName');
};
export default new SpeContainerListCommand();
//# sourceMappingURL=container-list.js.map