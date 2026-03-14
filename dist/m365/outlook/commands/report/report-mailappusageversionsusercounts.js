import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class OutlookReportMailAppUsageVersionsUserCountsCommand extends PeriodBasedReport {
    get name() {
        return commands.REPORT_MAILAPPUSAGEVERSIONSUSERCOUNTS;
    }
    get usageEndpoint() {
        return 'getEmailAppUsageVersionsUserCounts';
    }
    get description() {
        return 'Gets the count of unique users by Outlook desktop version.';
    }
}
export default new OutlookReportMailAppUsageVersionsUserCountsCommand();
//# sourceMappingURL=report-mailappusageversionsusercounts.js.map