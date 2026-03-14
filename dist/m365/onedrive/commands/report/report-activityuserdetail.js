import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';
class OneDriveReportActivityUserDetailCommand extends DateAndPeriodBasedReport {
    get name() {
        return commands.REPORT_ACTIVITYUSERDETAIL;
    }
    get usageEndpoint() {
        return 'getOneDriveActivityUserDetail';
    }
    get description() {
        return 'Gets details about OneDrive activity by user';
    }
}
export default new OneDriveReportActivityUserDetailCommand();
//# sourceMappingURL=report-activityuserdetail.js.map