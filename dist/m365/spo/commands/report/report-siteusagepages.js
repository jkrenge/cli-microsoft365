import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class SpoReportSiteUsagePagesCommand extends PeriodBasedReport {
    get name() {
        return commands.REPORT_SITEUSAGEPAGES;
    }
    get usageEndpoint() {
        return 'getSharePointSiteUsagePages';
    }
    get description() {
        return 'Gets the number of pages viewed across all sites';
    }
}
export default new SpoReportSiteUsagePagesCommand();
//# sourceMappingURL=report-siteusagepages.js.map