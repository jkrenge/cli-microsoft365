import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class OneDriveReportActivityUserCountCommand extends PeriodBasedReport {
    get name() {
        return commands.REPORT_ACTIVITYUSERCOUNTS;
    }
    get usageEndpoint() {
        return 'getOneDriveActivityUserCounts';
    }
    get description() {
        return 'Gets the trend in the number of active OneDrive users';
    }
}
export default new OneDriveReportActivityUserCountCommand();
//# sourceMappingURL=report-activityusercounts.js.map