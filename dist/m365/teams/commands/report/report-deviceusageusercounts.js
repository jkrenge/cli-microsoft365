import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class TeamsReportDeviceUsageUserCountsCommand extends PeriodBasedReport {
    get name() {
        return commands.REPORT_DEVICEUSAGEUSERCOUNTS;
    }
    get description() {
        return 'Get the number of Microsoft Teams daily unique users by device type';
    }
    get usageEndpoint() {
        return 'getTeamsDeviceUsageUserCounts';
    }
}
export default new TeamsReportDeviceUsageUserCountsCommand();
//# sourceMappingURL=report-deviceusageusercounts.js.map