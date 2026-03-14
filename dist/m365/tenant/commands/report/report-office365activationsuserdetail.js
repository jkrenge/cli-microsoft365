import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TenantReportOffice365ActivationsUserDetailCommand extends GraphCommand {
    get allowedOutputs() {
        return ['json', 'csv'];
    }
    get name() {
        return commands.REPORT_OFFICE365ACTIVATIONSUSERDETAIL;
    }
    get description() {
        return 'Get details about users who have activated Microsoft 365.';
    }
    async commandAction(logger, args) {
        const endpoint = `${this.resource}/v1.0/reports/getOffice365ActivationsUserDetail`;
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
export default new TenantReportOffice365ActivationsUserDetailCommand();
//# sourceMappingURL=report-office365activationsuserdetail.js.map