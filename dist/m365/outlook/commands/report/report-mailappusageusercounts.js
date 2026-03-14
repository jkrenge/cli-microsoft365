import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class OutlookReportMailAppUsageUserCountsCommand extends PeriodBasedReport {
    get name() {
        return commands.REPORT_MAILAPPUSAGEUSERCOUNTS;
    }
    get usageEndpoint() {
        return 'getEmailAppUsageUserCounts';
    }
    get description() {
        return 'Gets the count of unique users that connected to Exchange Online using any email app';
    }
}
export default new OutlookReportMailAppUsageUserCountsCommand();
//# sourceMappingURL=report-mailappusageusercounts.js.map