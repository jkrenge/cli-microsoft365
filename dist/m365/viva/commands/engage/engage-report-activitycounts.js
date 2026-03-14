import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class VivaEngageReportActivityCountsCommand extends PeriodBasedReport {
    get name() {
        return commands.ENGAGE_REPORT_ACTIVITYCOUNTS;
    }
    get usageEndpoint() {
        return 'getYammerActivityCounts';
    }
    get description() {
        return 'Gets the trends on the amount of VivaEngage activity in your organization by how many messages were posted, read, and liked';
    }
}
export default new VivaEngageReportActivityCountsCommand();
//# sourceMappingURL=engage-report-activitycounts.js.map