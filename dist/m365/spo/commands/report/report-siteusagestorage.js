import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class SpoReportSiteUsageStorageCommand extends PeriodBasedReport {
    get name() {
        return commands.REPORT_SITEUSAGESTORAGE;
    }
    get usageEndpoint() {
        return 'getSharePointSiteUsageStorage';
    }
    get description() {
        return 'Gets the trend of storage allocated and consumed during the reporting period';
    }
}
export default new SpoReportSiteUsageStorageCommand();
//# sourceMappingURL=report-siteusagestorage.js.map