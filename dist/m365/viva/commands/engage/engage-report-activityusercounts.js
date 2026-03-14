import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class VivaEngageReportActivityUserCountsCommand extends PeriodBasedReport {
    get name() {
        return commands.ENGAGE_REPORT_ACTIVITYUSERCOUNTS;
    }
    get usageEndpoint() {
        return 'getYammerActivityUserCounts';
    }
    get description() {
        return 'Gets the trends on the number of unique users who posted, read, and liked Viva Engage messages';
    }
}
export default new VivaEngageReportActivityUserCountsCommand();
//# sourceMappingURL=engage-report-activityusercounts.js.map