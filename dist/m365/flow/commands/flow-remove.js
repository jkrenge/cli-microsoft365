var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _FlowRemoveCommand_instances, _FlowRemoveCommand_initTelemetry, _FlowRemoveCommand_initOptions, _FlowRemoveCommand_initValidators;
import { cli } from '../../../cli/cli.js';
import request from '../../../request.js';
import { formatting } from '../../../utils/formatting.js';
import { validation } from '../../../utils/validation.js';
import commands from '../commands.js';
import PowerAutomateCommand from '../../base/PowerAutomateCommand.js';
class FlowRemoveCommand extends PowerAutomateCommand {
    get name() {
        return commands.REMOVE;
    }
    get description() {
        return 'Removes the specified Microsoft Flow';
    }
    constructor() {
        super();
        _FlowRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _FlowRemoveCommand_instances, "m", _FlowRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _FlowRemoveCommand_instances, "m", _FlowRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _FlowRemoveCommand_instances, "m", _FlowRemoveCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Removing Microsoft Flow ${args.options.name}...`);
        }
        const removeFlow = async () => {
            const requestOptions = {
                url: `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/${args.options.asAdmin ? 'scopes/admin/' : ''}environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.name)}?api-version=2016-11-01`,
                fullResponse: true,
                headers: {
                    accept: 'application/json'
                },
                responseType: 'json'
            };
            try {
                const rawRes = await request.delete(requestOptions);
                // handle 204 and throw error message to cmd when invalid flow id is passed
                // https://github.com/pnp/cli-microsoft365/issues/1063#issuecomment-537218957
                if (rawRes.statusCode === 204) {
                    throw `Error: Resource '${args.options.name}' does not exist in environment '${args.options.environmentName}'`;
                }
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeFlow();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the Microsoft Flow ${args.options.name}?` });
            if (result) {
                await removeFlow();
            }
        }
    }
}
_FlowRemoveCommand_instances = new WeakSet(), _FlowRemoveCommand_initTelemetry = function _FlowRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            asAdmin: !!args.options.asAdmin,
            force: !!args.options.force
        });
    });
}, _FlowRemoveCommand_initOptions = function _FlowRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name <name>'
    }, {
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '--asAdmin'
    }, {
        option: '-f, --force'
    });
}, _FlowRemoveCommand_initValidators = function _FlowRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.name)) {
            return `${args.options.name} is not a valid GUID`;
        }
        return true;
    });
};
export default new FlowRemoveCommand();
//# sourceMappingURL=flow-remove.js.map