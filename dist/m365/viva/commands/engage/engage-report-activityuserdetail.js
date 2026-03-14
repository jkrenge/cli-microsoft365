import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';
class VivaEngageReportActivityUserDetailCommand extends DateAndPeriodBasedReport {
    get name() {
        return commands.ENGAGE_REPORT_ACTIVITYUSERDETAIL;
    }
    get usageEndpoint() {
        return 'getYammerActivityUserDetail';
    }
    get description() {
        return 'Gets details about Viva Engage activity by user';
    }
}
export default new VivaEngageReportActivityUserDetailCommand();
//# sourceMappingURL=engage-report-activityuserdetail.js.map