import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class OutlookReportMailActivityUserCountsCommand extends PeriodBasedReport {
    get name() {
        return commands.REPORT_MAILACTIVITYUSERCOUNTS;
    }
    get usageEndpoint() {
        return 'getEmailActivityUserCounts';
    }
    get description() {
        return 'Enables you to understand trends on the number of unique users who are performing email activities like send, read, and receive';
    }
}
export default new OutlookReportMailActivityUserCountsCommand();
//# sourceMappingURL=report-mailactivityusercounts.js.map