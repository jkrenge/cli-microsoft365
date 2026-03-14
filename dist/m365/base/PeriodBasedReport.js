var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PeriodBasedReport_instances, _PeriodBasedReport_initOptions, _PeriodBasedReport_initValidators;
import request from '../../request.js';
import { formatting } from '../../utils/formatting.js';
import GraphCommand from "./GraphCommand.js";
class PeriodBasedReport extends GraphCommand {
    get allowedOutputs() {
        return ['json', 'csv'];
    }
    constructor() {
        super();
        _PeriodBasedReport_instances.add(this);
        __classPrivateFieldGet(this, _PeriodBasedReport_instances, "m", _PeriodBasedReport_initOptions).call(this);
        __classPrivateFieldGet(this, _PeriodBasedReport_instances, "m", _PeriodBasedReport_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const endpoint = `${this.resource}/v1.0/reports/${this.usageEndpoint}(period='${formatting.encodeQueryParameter(args.options.period)}')`;
        await this.executeReport(endpoint, logger, args.options.output);
    }
    async executeReport(endPoint, logger, output) {
        const requestOptions = {
            url: endPoint,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        let res;
        try {
            res = await request.get(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
            return;
        }
        let content;
        const cleanResponse = this.removeEmptyLines(res);
        if (output && output.toLowerCase() === 'json') {
            const reportData = this.getReport(cleanResponse);
            content = reportData;
        }
        else {
            content = cleanResponse;
        }
        await logger.log(content);
    }
    removeEmptyLines(input) {
        const rows = input.split('\n');
        const cleanRows = rows.filter(Boolean);
        return cleanRows.join('\n');
    }
    getReport(res) {
        const rows = res.split('\n');
        const jsonObj = [];
        const headers = rows[0].split(',');
        for (let i = 1; i < rows.length; i++) {
            const data = rows[i].split(',');
            const obj = {};
            for (let j = 0; j < data.length; j++) {
                obj[headers[j].trim()] = data[j].trim();
            }
            jsonObj.push(obj);
        }
        return jsonObj;
    }
    async validatePeriod(args) {
        const period = args.options.period;
        if (period &&
            ['D7', 'D30', 'D90', 'D180'].indexOf(period) < 0) {
            return `${period} is not a valid period type. The supported values are D7|D30|D90|D180`;
        }
        return true;
    }
}
_PeriodBasedReport_instances = new WeakSet(), _PeriodBasedReport_initOptions = function _PeriodBasedReport_initOptions() {
    this.options.push({
        option: '-p, --period <period>',
        autocomplete: ['D7', 'D30', 'D90', 'D180']
    });
}, _PeriodBasedReport_initValidators = function _PeriodBasedReport_initValidators() {
    this.validators.push((args) => this.validatePeriod(args));
};
export default PeriodBasedReport;
//# sourceMappingURL=PeriodBasedReport.js.map