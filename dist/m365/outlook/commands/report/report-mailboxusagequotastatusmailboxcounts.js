import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class OutlookReportMailboxUsageQuotaStatusMailboxCountsCommand extends PeriodBasedReport {
    get name() {
        return commands.REPORT_MAILBOXUSAGEQUOTASTATUSMAILBOXCOUNTS;
    }
    get usageEndpoint() {
        return 'getMailboxUsageQuotaStatusMailboxCounts';
    }
    get description() {
        return 'Gets the count of user mailboxes in each quota category';
    }
}
export default new OutlookReportMailboxUsageQuotaStatusMailboxCountsCommand();
//# sourceMappingURL=report-mailboxusagequotastatusmailboxcounts.js.map