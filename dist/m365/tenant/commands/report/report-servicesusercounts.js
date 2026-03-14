import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class TenantReportServicesUserCountsCommand extends PeriodBasedReport {
    get name() {
        return commands.REPORT_SERVICESUSERCOUNTS;
    }
    get description() {
        return 'Gets the count of users by activity type and service.';
    }
    get usageEndpoint() {
        return 'getOffice365ServicesUserCounts';
    }
}
export default new TenantReportServicesUserCountsCommand();
//# sourceMappingURL=report-servicesusercounts.js.map