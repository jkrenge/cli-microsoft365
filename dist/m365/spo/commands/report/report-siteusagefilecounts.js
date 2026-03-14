import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class SpoReportSiteUsageFileCountsCommand extends PeriodBasedReport {
    get name() {
        return commands.REPORT_SITEUSAGEFILECOUNTS;
    }
    get usageEndpoint() {
        return 'getSharePointSiteUsageFileCounts';
    }
    get description() {
        return 'Get the total number of files across all sites and the number of active files';
    }
}
export default new SpoReportSiteUsageFileCountsCommand();
//# sourceMappingURL=report-siteusagefilecounts.js.map