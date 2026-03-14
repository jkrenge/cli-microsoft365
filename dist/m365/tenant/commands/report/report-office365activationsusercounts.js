import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TenantReportOffice365ActivationsUserCountsCommand extends GraphCommand {
    get allowedOutputs() {
        return ['json', 'csv'];
    }
    get name() {
        return commands.REPORT_OFFICE365ACTIVATIONSUSERCOUNTS;
    }
    get description() {
        return 'Get the count of enabled users with activated Office subscriptions.';
    }
    async commandAction(logger, args) {
        const endpoint = `${this.resource}/v1.0/reports/getOffice365ActivationsUserCounts`;
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
export default new TenantReportOffice365ActivationsUserCountsCommand();
//# sourceMappingURL=report-office365activationsusercounts.js.map