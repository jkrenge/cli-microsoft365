import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TenantReportOffice365ActivationCountsCommand extends GraphCommand {
    get allowedOutputs() {
        return ['json', 'csv'];
    }
    get name() {
        return commands.REPORT_OFFICE365ACTIVATIONCOUNTS;
    }
    get description() {
        return 'Get the count of Microsoft 365 activations on desktops and devices.';
    }
    async commandAction(logger, args) {
        const endpoint = `${this.resource}/v1.0/reports/getOffice365ActivationCounts`;
        await this.loadReport(endpoint, logger, args.options.output);
    }
    async loadReport(endPoint, logger, output) {
        const requestOptions = {
            url: endPoint,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        try {
            const res = await request.get(requestOptions);
            let content = '';
            if (output && output.toLowerCase() === 'json') {
                content = formatting.parseCsvToJson(res);
            }
            else {
                content = res;
            }
            await logger.log(content);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new TenantReportOffice365ActivationCountsCommand();
//# sourceMappingURL=report-office365activationcounts.js.map