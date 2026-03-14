var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PlannerRosterRemoveCommand_instances, _PlannerRosterRemoveCommand_initTelemetry, _PlannerRosterRemoveCommand_initOptions, _PlannerRosterRemoveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PlannerRosterRemoveCommand extends GraphCommand {
    get name() {
        return commands.ROSTER_REMOVE;
    }
    get description() {
        return 'Removes a Microsoft Planner Roster';
    }
    constructor() {
        super();
        _PlannerRosterRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _PlannerRosterRemoveCommand_instances, "m", _PlannerRosterRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PlannerRosterRemoveCommand_instances, "m", _PlannerRosterRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PlannerRosterRemoveCommand_instances, "m", _PlannerRosterRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeRoster(args, logger);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove roster ${args.options.id}?` });
            if (result) {
                await this.removeRoster(args, logger);
            }
        }
    }
    async removeRoster(args, logger) {
        if (this.verbose) {
            await logger.logToStderr(`Removing roster ${args.options.id}`);
        }
        try {
            const requestOptions = {
                url: `${this.resource}/beta/planner/rosters/${args.options.id}`,
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
_PlannerRosterRemoveCommand_instances = new WeakSet(), _PlannerRosterRemoveCommand_initTelemetry = function _PlannerRosterRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: !!args.options.force
        });
    });
}, _PlannerRosterRemoveCommand_initOptions = function _PlannerRosterRemoveCommand_initOptions() {
    this.options.unshift({
        option: '--id <id>'
    }, {
        option: '-f, --force'
    });
}, _PlannerRosterRemoveCommand_initTypes = function _PlannerRosterRemoveCommand_initTypes() {
    this.types.string.push('id');
};
export default new PlannerRosterRemoveCommand();
//# sourceMappingURL=roster-remove.js.map