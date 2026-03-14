import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';
class SpoReportActivityUserDetailCommand extends DateAndPeriodBasedReport {
    get name() {
        return commands.REPORT_ACTIVITYUSERDETAIL;
    }
    get description() {
        return 'Gets details about SharePoint activity by user';
    }
    get usageEndpoint() {
        return 'getSharePointActivityUserDetail';
    }
}
export default new SpoReportActivityUserDetailCommand();
//# sourceMappingURL=report-activityuserdetail.js.map