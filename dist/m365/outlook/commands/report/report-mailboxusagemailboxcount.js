import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class OutlookReportMailboxUsageMailboxCountCommand extends PeriodBasedReport {
    get name() {
        return commands.REPORT_MAILBOXUSAGEMAILBOXCOUNT;
    }
    get usageEndpoint() {
        return 'getMailboxUsageMailboxCounts';
    }
    get description() {
        return 'Gets the total number of user mailboxes in your organization and how many are active each day of the reporting period.';
    }
}
export default new OutlookReportMailboxUsageMailboxCountCommand();
//# sourceMappingURL=report-mailboxusagemailboxcount.js.map