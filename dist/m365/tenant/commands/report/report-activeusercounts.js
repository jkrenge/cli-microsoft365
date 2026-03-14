import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class TenantReportActiveUserCountsCommand extends PeriodBasedReport {
    get name() {
        return commands.REPORT_ACTIVEUSERCOUNTS;
    }
    get description() {
        return 'Gets the count of daily active users in the reporting period by product.';
    }
    get usageEndpoint() {
        return 'getOffice365ActiveUserCounts';
    }
}
export default new TenantReportActiveUserCountsCommand();
//# sourceMappingURL=report-activeusercounts.js.map