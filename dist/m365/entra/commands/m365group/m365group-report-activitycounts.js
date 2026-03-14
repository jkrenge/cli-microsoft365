import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class M365GroupReportActivityCountsCommand extends PeriodBasedReport {
    get name() {
        return commands.M365GROUP_REPORT_ACTIVITYCOUNTS;
    }
    get description() {
        return 'Get the number of group activities across group workloads';
    }
    get usageEndpoint() {
        return 'getOffice365GroupsActivityCounts';
    }
}
export default new M365GroupReportActivityCountsCommand();
//# sourceMappingURL=m365group-report-activitycounts.js.map