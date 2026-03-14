import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';
class OutlookReportMailActivityUserDetailCommand extends DateAndPeriodBasedReport {
    get name() {
        return commands.REPORT_MAILACTIVITYUSERDETAIL;
    }
    get usageEndpoint() {
        return 'getEmailActivityUserDetail';
    }
    get description() {
        return 'Gets details about email activity users have performed';
    }
}
export default new OutlookReportMailActivityUserDetailCommand();
//# sourceMappingURL=report-mailactivityuserdetail.js.map