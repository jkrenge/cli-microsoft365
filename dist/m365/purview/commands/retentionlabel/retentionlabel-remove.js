var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PurviewRetentionLabelRemoveCommand_instances, _PurviewRetentionLabelRemoveCommand_initTelemetry, _PurviewRetentionLabelRemoveCommand_initOptions, _PurviewRetentionLabelRemoveCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PurviewRetentionLabelRemoveCommand extends GraphCommand {
    get name() {
        return commands.RETENTIONLABEL_REMOVE;
    }
    get description() {
        return 'Delete a retention label';
    }
    constructor() {
        super();
        _PurviewRetentionLabelRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _PurviewRetentionLabelRemoveCommand_instances, "m", _PurviewRetentionLabelRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PurviewRetentionLabelRemoveCommand_instances, "m", _PurviewRetentionLabelRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PurviewRetentionLabelRemoveCommand_instances, "m", _PurviewRetentionLabelRemoveCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeRetentionLabel(args);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the retention label ${args.options.id}?` });
            if (result) {
                await this.removeRetentionLabel(args);
            }
        }
    }
    async removeRetentionLabel(args) {
        try {
            const requestOptions = {
                url: `${this.resource}/beta/security/labels/retentionLabels/${args.options.id}`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            await request.delete(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_PurviewRetentionLabelRemoveCommand_instances = new WeakSet(), _PurviewRetentionLabelRemoveCommand_initTelemetry = function _PurviewRetentionLabelRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: !!args.options.force
        });
    });
}, _PurviewRetentionLabelRemoveCommand_initOptions = function _PurviewRetentionLabelRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '-f, --force'
    });
}, _PurviewRetentionLabelRemoveCommand_initValidators = function _PurviewRetentionLabelRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.id)) {
            return `'${args.options.id}' is not a valid GUID.`;
        }
        return true;
    });
};
export default new PurviewRetentionLabelRemoveCommand();
//# sourceMappingURL=retentionlabel-remove.js.map