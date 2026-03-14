import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';
class SpoReportSiteUsageDetailCommand extends DateAndPeriodBasedReport {
    get name() {
        return commands.REPORT_SITEUSAGEDETAIL;
    }
    get description() {
        return 'Gets details about SharePoint site usage';
    }
    get usageEndpoint() {
        return 'getSharePointSiteUsageDetail';
    }
}
export default new SpoReportSiteUsageDetailCommand();
//# sourceMappingURL=report-siteusagedetail.js.map