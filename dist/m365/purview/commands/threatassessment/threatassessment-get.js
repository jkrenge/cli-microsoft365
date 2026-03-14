var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PurviewThreatAssessmentGetCommand_instances, _PurviewThreatAssessmentGetCommand_initTelemetry, _PurviewThreatAssessmentGetCommand_initOptions, _PurviewThreatAssessmentGetCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PurviewThreatAssessmentGetCommand extends GraphCommand {
    get name() {
        return commands.THREATASSESSMENT_GET;
    }
    get description() {
        return 'Get a threat assessment';
    }
    constructor() {
        super();
        _PurviewThreatAssessmentGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _PurviewThreatAssessmentGetCommand_instances, "m", _PurviewThreatAssessmentGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PurviewThreatAssessmentGetCommand_instances, "m", _PurviewThreatAssessmentGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PurviewThreatAssessmentGetCommand_instances, "m", _PurviewThreatAssessmentGetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Retrieving threat assessment with id ${args.options.id}`);
            }
            const requestOptions = {
                url: `${this.resource}/v1.0/informationProtection/threatAssessmentRequests/${args.options.id}${args.options.withResults ? '?$expand=results' : ''}`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            const res = await request.get(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_PurviewThreatAssessmentGetCommand_instances = new WeakSet(), _PurviewThreatAssessmentGetCommand_initTelemetry = function _PurviewThreatAssessmentGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            includeResults: !!args.options.includeResults,
            withResults: !!args.options.withResults
        });
    });
}, _PurviewThreatAssessmentGetCommand_initOptions = function _PurviewThreatAssessmentGetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '--withResults'
    });
}, _PurviewThreatAssessmentGetCommand_initValidators = function _PurviewThreatAssessmentGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID.`;
        }
        return true;
    });
};
export default new PurviewThreatAssessmentGetCommand();
//# sourceMappingURL=threatassessment-get.js.map