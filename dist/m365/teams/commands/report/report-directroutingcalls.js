var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsReportDirectroutingcallsCommand_instances, _TeamsReportDirectroutingcallsCommand_initTelemetry, _TeamsReportDirectroutingcallsCommand_initOptions, _TeamsReportDirectroutingcallsCommand_initValidators;
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphApplicationCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsReportDirectroutingcallsCommand extends GraphApplicationCommand {
    get allowedOutputs() {
        return ['json', 'csv'];
    }
    get name() {
        return commands.REPORT_DIRECTROUTINGCALLS;
    }
    get description() {
        return 'Get details about direct routing calls made within a given time period';
    }
    defaultProperties() {
        return ['id', 'calleeNumber', 'callerNumber', 'startDateTime'];
    }
    constructor() {
        super();
        _TeamsReportDirectroutingcallsCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsReportDirectroutingcallsCommand_instances, "m", _TeamsReportDirectroutingcallsCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsReportDirectroutingcallsCommand_instances, "m", _TeamsReportDirectroutingcallsCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsReportDirectroutingcallsCommand_instances, "m", _TeamsReportDirectroutingcallsCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const toDateTimeParameter = formatting.encodeQueryParameter(args.options.toDateTime ? args.options.toDateTime : new Date().toISOString());
        try {
            const response = await odata.getAllItems(`${this.resource}/v1.0/communications/callRecords/getDirectRoutingCalls(fromDateTime=${formatting.encodeQueryParameter(args.options.fromDateTime)},toDateTime=${toDateTimeParameter})`);
            await logger.log(response);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_TeamsReportDirectroutingcallsCommand_instances = new WeakSet(), _TeamsReportDirectroutingcallsCommand_initTelemetry = function _TeamsReportDirectroutingcallsCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            toDateTime: typeof args.options.toDateTime !== 'undefined'
        });
    });
}, _TeamsReportDirectroutingcallsCommand_initOptions = function _TeamsReportDirectroutingcallsCommand_initOptions() {
    this.options.unshift({
        option: '--fromDateTime <fromDateTime>'
    }, {
        option: '--toDateTime [toDateTime]'
    });
}, _TeamsReportDirectroutingcallsCommand_initValidators = function _TeamsReportDirectroutingcallsCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidISODateDashOnly(args.options.fromDateTime)) {
            return 'The fromDateTime is not a valid ISO date string (YYYY-MM-DD)';
        }
        if (args.options.toDateTime &&
            !validation.isValidISODateDashOnly(args.options.toDateTime)) {
            return 'The toDateTime is not a valid ISO date string (YYYY-MM-DD)';
        }
        if (Math.ceil((new Date(args.options.toDateTime || new Date().toISOString()).getTime() - new Date(args.options.fromDateTime).getTime()) / (1000 * 3600 * 24)) > 90) {
            return 'The maximum number of days between fromDateTime and toDateTime cannot exceed 90';
        }
        return true;
    });
};
export default new TeamsReportDirectroutingcallsCommand();
//# sourceMappingURL=report-directroutingcalls.js.map