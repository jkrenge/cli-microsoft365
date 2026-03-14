import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class SpoReportActivityUserCountsCommand extends PeriodBasedReport {
    get name() {
        return commands.REPORT_ACTIVITYUSERCOUNTS;
    }
    get usageEndpoint() {
        return 'getSharePointActivityUserCounts';
    }
    get description() {
        return 'Gets the trend in the number of active users';
    }
}
export default new SpoReportActivityUserCountsCommand();
//# sourceMappingURL=report-activityusercounts.js.map