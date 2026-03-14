import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class VivaEngageReportDeviceUsageUserCountsCommand extends PeriodBasedReport {
    get name() {
        return commands.ENGAGE_REPORT_DEVICEUSAGEUSERCOUNTS;
    }
    get usageEndpoint() {
        return 'getYammerDeviceUsageUserCounts';
    }
    get description() {
        return 'Gets the number of daily users by device type';
    }
}
export default new VivaEngageReportDeviceUsageUserCountsCommand();
//# sourceMappingURL=engage-report-deviceusageusercounts.js.map