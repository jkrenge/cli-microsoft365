var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _DateAndPeriodBasedReport_instances, _DateAndPeriodBasedReport_initTelemetry, _DateAndPeriodBasedReport_initOptions, _DateAndPeriodBasedReport_initValidators;
import { formatting } from '../../utils/formatting.js';
import PeriodBasedReport from './PeriodBasedReport.js';
class DateAndPeriodBasedReport extends PeriodBasedReport {
    constructor() {
        super();
        _DateAndPeriodBasedReport_instances.add(this);
        __classPrivateFieldGet(this, _DateAndPeriodBasedReport_instances, "m", _DateAndPeriodBasedReport_initTelemetry).call(this);
        __classPrivateFieldGet(this, _DateAndPeriodBasedReport_instances, "m", _DateAndPeriodBasedReport_initOptions).call(this);
        __classPrivateFieldGet(this, _DateAndPeriodBasedReport_instances, "m", _DateAndPeriodBasedReport_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const periodParameter = args.options.period ? `${this.usageEndpoint}(period='${formatting.encodeQueryParameter(args.options.period)}')` : '';
        const dateParameter = args.options.date ? `${this.usageEndpoint}(date=${formatting.encodeQueryParameter(args.options.date)})` : '';
        const endpoint = `${this.resource}/v1.0/reports/${(args.options.period ? periodParameter : dateParameter)}`;
        await this.executeReport(endpoint, logger, args.options.output);
    }
}
_DateAndPeriodBasedReport_instances = new WeakSet(), _DateAndPeriodBasedReport_initTelemetry = function _DateAndPeriodBasedReport_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            period: args.options.period,
            date: typeof args.options.date !== 'undefined'
        });
    });
}, _DateAndPeriodBasedReport_initOptions = function _DateAndPeriodBasedReport_initOptions() {
    this.options.unshift({ option: '-d, --date [date]' });
    this.options.forEach(option => {
        option.option = option.option.replace('-p, --period <period>', '-p, --period [period]');
    });
}, _DateAndPeriodBasedReport_initValidators = function _DateAndPeriodBasedReport_initValidators() {
    this.validators.push(async (args) => {
        if (!args.options.period && !args.options.date) {
            return 'Specify period or date, one is required.';
        }
        if (args.options.period && args.options.date) {
            return 'Specify period or date but not both.';
        }
        if (args.options.date && !(args.options.date.match(/^\d{4}-\d{2}-\d{2}$/))) {
            return `${args.options.date} is not a valid date. The supported date format is YYYY-MM-DD`;
        }
        return true;
    });
};
export default DateAndPeriodBasedReport;
//# sourceMappingURL=DateAndPeriodBasedReport.js.map