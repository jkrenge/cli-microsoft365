import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class OutlookReportMailboxUsageDetailCommand extends PeriodBasedReport {
    get name() {
        return commands.REPORT_MAILBOXUSAGEDETAIL;
    }
    get usageEndpoint() {
        return 'getMailboxUsageDetail';
    }
    get description() {
        return 'Gets details about mailbox usage';
    }
}
export default new OutlookReportMailboxUsageDetailCommand();
//# sourceMappingURL=report-mailboxusagedetail.js.map