import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';
class TeamsReportUserActivityUserDetailCommand extends DateAndPeriodBasedReport {
    get name() {
        return commands.REPORT_USERACTIVITYUSERDETAIL;
    }
    get usageEndpoint() {
        return 'getTeamsUserActivityUserDetail';
    }
    get description() {
        return 'Get details about Microsoft Teams user activity by user.';
    }
}
export default new TeamsReportUserActivityUserDetailCommand();
//# sourceMappingURL=report-useractivityuserdetail.js.map