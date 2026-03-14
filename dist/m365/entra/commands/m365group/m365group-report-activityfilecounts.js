import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class M365GroupReportActivityFileCountsCommand extends PeriodBasedReport {
    get name() {
        return commands.M365GROUP_REPORT_ACTIVITYFILECOUNTS;
    }
    get description() {
        return 'Get the total number of files and how many of them were active across all group sites associated with an Microsoft 365 Group';
    }
    get usageEndpoint() {
        return 'getOffice365GroupsActivityFileCounts';
    }
}
export default new M365GroupReportActivityFileCountsCommand();
//# sourceMappingURL=m365group-report-activityfilecounts.js.map