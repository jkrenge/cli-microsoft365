import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class TeamsReportDeviceUsageDistributionUserCountsCommand extends PeriodBasedReport {
    get name() {
        return commands.REPORT_DEVICEUSAGEDISTRIBUTIONUSERCOUNTS;
    }
    get usageEndpoint() {
        return 'getTeamsDeviceUsageDistributionUserCounts';
    }
    get description() {
        return 'Get the number of Microsoft Teams unique users by device type';
    }
}
export default new TeamsReportDeviceUsageDistributionUserCountsCommand();
//# sourceMappingURL=report-deviceusagedistributionusercounts.js.map