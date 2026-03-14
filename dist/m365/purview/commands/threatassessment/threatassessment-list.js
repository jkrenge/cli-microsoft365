var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PurviewThreatAssessmentListCommand_instances, _a, _PurviewThreatAssessmentListCommand_initTelemetry, _PurviewThreatAssessmentListCommand_initOptions, _PurviewThreatAssessmentListCommand_initValidators;
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PurviewThreatAssessmentListCommand extends GraphCommand {
    get name() {
        return commands.THREATASSESSMENT_LIST;
    }
    get description() {
        return 'Get a list of threat assessments';
    }
    defaultProperties() {
        return ['id', 'type', 'category'];
    }
    constructor() {
        super();
        _PurviewThreatAssessmentListCommand_instances.add(this);
        __classPrivateFieldGet(this, _PurviewThreatAssessmentListCommand_instances, "m", _PurviewThreatAssessmentListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PurviewThreatAssessmentListCommand_instances, "m", _PurviewThreatAssessmentListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PurviewThreatAssessmentListCommand_instances, "m", _PurviewThreatAssessmentListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr('Retrieving a list of threat assessments');
        }
        try {
            const filter = this.getFilterQuery(args.options);
            const items = await odata.getAllItems(`${this.resource}/v1.0/informationProtection/threatAssessmentRequests${filter}`, 'minimal');
            let itemsToReturn = [];
            switch (args.options.type) {
                case 'mail':
                    itemsToReturn = items.filter(item => item['@odata.type'] === '#microsoft.graph.mailAssessmentRequest');
                    break;
                case 'emailFile':
                    itemsToReturn = items.filter(item => item['@odata.type'] === '#microsoft.graph.emailFileAssessmentRequest');
                    break;
                default:
                    itemsToReturn = items;
                    break;
            }
            for (const item of itemsToReturn) {
                item['type'] = this.getConvertedType(item['@odata.type']);
                delete item['@odata.type'];
            }
            await logger.log(itemsToReturn);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    // Content type is not equal to type. 
    // Threat assessments of type emailFile have contentType mail as well.
    // This function gets the correct filter URL to be able to query the least amount of data
    getFilterQuery(options) {
        if (options.type === undefined) {
            return '';
        }
        if (options.type === 'emailFile') {
            return `?$filter=contentType eq 'mail'`;
        }
        return `?$filter=contentType eq '${options.type}'`;
    }
    getConvertedType(type) {
        switch (type) {
            case '#microsoft.graph.mailAssessmentRequest':
                return 'mail';
            case '#microsoft.graph.fileAssessmentRequest':
                return 'file';
            case '#microsoft.graph.emailFileAssessmentRequest':
                return 'emailFile';
            case '#microsoft.graph.urlAssessmentRequest':
                return 'url';
            default:
                return 'Unknown';
        }
    }
}
_a = PurviewThreatAssessmentListCommand, _PurviewThreatAssessmentListCommand_instances = new WeakSet(), _PurviewThreatAssessmentListCommand_initTelemetry = function _PurviewThreatAssessmentListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            type: typeof args.options.type !== 'undefined'
        });
    });
}, _PurviewThreatAssessmentListCommand_initOptions = function _PurviewThreatAssessmentListCommand_initOptions() {
    this.options.unshift({
        option: '-t, --type [type]',
        autocomplete: _a.allowedTypes
    });
}, _PurviewThreatAssessmentListCommand_initValidators = function _PurviewThreatAssessmentListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.type && _a.allowedTypes.indexOf(args.options.type) < 0) {
            return `${args.options.type} is not a valid type. Allowed values are ${_a.allowedTypes.join(', ')}`;
        }
        return true;
    });
};
PurviewThreatAssessmentListCommand.allowedTypes = ['mail', 'file', 'emailFile', 'url'];
export default new PurviewThreatAssessmentListCommand();
//# sourceMappingURL=threatassessment-list.js.map