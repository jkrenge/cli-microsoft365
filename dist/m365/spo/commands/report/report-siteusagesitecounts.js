import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class SpoReportSiteUsageSiteCountsCommand extends PeriodBasedReport {
    get name() {
        return commands.REPORT_SITEUSAGESITECOUNTS;
    }
    get usageEndpoint() {
        return 'getSharePointSiteUsageSiteCounts';
    }
    get description() {
        return 'Gets the total number of files across all sites and the number of active files';
    }
}
export default new SpoReportSiteUsageSiteCountsCommand();
//# sourceMappingURL=report-siteusagesitecounts.js.map