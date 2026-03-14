import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class VivaEngageReportGroupsActivityGroupCountsCommand extends PeriodBasedReport {
    get name() {
        return commands.ENGAGE_REPORT_GROUPSACTIVITYGROUPCOUNTS;
    }
    get usageEndpoint() {
        return 'getYammerGroupsActivityGroupCounts';
    }
    get description() {
        return 'Gets the total number of groups that existed and how many included group conversation activity';
    }
}
export default new VivaEngageReportGroupsActivityGroupCountsCommand();
//# sourceMappingURL=engage-report-groupsactivitygroupcounts.js.map