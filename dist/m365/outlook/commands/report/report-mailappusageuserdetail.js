import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';
class OutlookReportMailAppUsageUserDetailCommand extends DateAndPeriodBasedReport {
    get name() {
        return commands.REPORT_MAILAPPUSAGEUSERDETAIL;
    }
    get usageEndpoint() {
        return 'getEmailAppUsageUserDetail';
    }
    get description() {
        return 'Gets details about which activities users performed on the various email apps';
    }
}
export default new OutlookReportMailAppUsageUserDetailCommand();
//# sourceMappingURL=report-mailappusageuserdetail.js.map