import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class OutlookReportMailActivityCountsCommand extends PeriodBasedReport {
    get name() {
        return commands.REPORT_MAILACTIVITYCOUNTS;
    }
    get usageEndpoint() {
        return 'getEmailActivityCounts';
    }
    get description() {
        return 'Enables you to understand the trends of email activity (like how many were sent, read, and received) in your organization';
    }
}
export default new OutlookReportMailActivityCountsCommand();
//# sourceMappingURL=report-mailactivitycounts.js.map