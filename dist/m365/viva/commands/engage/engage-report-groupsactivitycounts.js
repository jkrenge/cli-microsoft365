import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class VivaEngageReportGroupsActivityCountsCommand extends PeriodBasedReport {
    get name() {
        return commands.ENGAGE_REPORT_GROUPSACTIVITYCOUNTS;
    }
    get usageEndpoint() {
        return 'getYammerGroupsActivityCounts';
    }
    get description() {
        return 'Gets the number of Viva Engage messages posted, read, and liked in groups';
    }
}
export default new VivaEngageReportGroupsActivityCountsCommand();
//# sourceMappingURL=engage-report-groupsactivitycounts.js.map