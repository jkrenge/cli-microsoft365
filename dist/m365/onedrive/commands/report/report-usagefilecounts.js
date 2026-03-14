import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class OneDriveReportUsageFileCountsCommand extends PeriodBasedReport {
    get name() {
        return commands.REPORT_USAGEFILECOUNTS;
    }
    get usageEndpoint() {
        return 'getOneDriveUsageFileCounts';
    }
    get description() {
        return 'Gets the total number of files across all sites and how many are active files';
    }
}
export default new OneDriveReportUsageFileCountsCommand();
//# sourceMappingURL=report-usagefilecounts.js.map