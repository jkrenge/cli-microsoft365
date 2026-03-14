import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class OutlookReportMailAppUsageAppsUserCountsCommand extends PeriodBasedReport {
    get name() {
        return commands.REPORT_MAILAPPUSAGEAPPSUSERCOUNTS;
    }
    get usageEndpoint() {
        return 'getEmailAppUsageAppsUserCounts';
    }
    get description() {
        return 'Gets the count of unique users per email app';
    }
}
export default new OutlookReportMailAppUsageAppsUserCountsCommand();
//# sourceMappingURL=report-mailappusageappsusercounts.js.map