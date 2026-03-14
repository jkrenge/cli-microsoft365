import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class SpoReportActivityFileCountsCommand extends PeriodBasedReport {
    get name() {
        return commands.REPORT_ACTIVITYFILECOUNTS;
    }
    get usageEndpoint() {
        return 'getSharePointActivityFileCounts';
    }
    get description() {
        return 'Gets the number of unique, licensed users who interacted with files stored on SharePoint sites';
    }
}
export default new SpoReportActivityFileCountsCommand();
//# sourceMappingURL=report-activityfilecounts.js.map