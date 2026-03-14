import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class VivaEngageReportDeviceUsageDistributionUserCountsCommand extends PeriodBasedReport {
    get name() {
        return commands.ENGAGE_REPORT_DEVICEUSAGEDISTRIBUTIONUSERCOUNTS;
    }
    get usageEndpoint() {
        return 'getYammerDeviceUsageDistributionUserCounts';
    }
    get description() {
        return 'Gets the number of users by device type';
    }
}
export default new VivaEngageReportDeviceUsageDistributionUserCountsCommand();
//# sourceMappingURL=engage-report-deviceusagedistributionusercounts.js.map