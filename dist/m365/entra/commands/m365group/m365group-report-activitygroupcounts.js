import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class M365GroupReportActivityGroupCountsCommand extends PeriodBasedReport {
    get name() {
        return commands.M365GROUP_REPORT_ACTIVITYGROUPCOUNTS;
    }
    get description() {
        return 'Get the daily total number of groups and how many of them were active based on email conversations, Viva Engage posts, and SharePoint file activities';
    }
    get usageEndpoint() {
        return 'getOffice365GroupsActivityGroupCounts';
    }
}
export default new M365GroupReportActivityGroupCountsCommand();
//# sourceMappingURL=m365group-report-activitygroupcounts.js.map