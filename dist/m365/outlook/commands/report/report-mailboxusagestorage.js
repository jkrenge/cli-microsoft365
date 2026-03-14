import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class OutlookReportMailboxUsageStorageCommand extends PeriodBasedReport {
    get name() {
        return commands.REPORT_MAILBOXUSAGESTORAGE;
    }
    get usageEndpoint() {
        return 'getMailboxUsageStorage';
    }
    get description() {
        return 'Gets the amount of mailbox storage used in your organization';
    }
}
export default new OutlookReportMailboxUsageStorageCommand();
//# sourceMappingURL=report-mailboxusagestorage.js.map