import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';
class TeamsReportDeviceUsageUserDetailCommand extends DateAndPeriodBasedReport {
    get name() {
        return commands.REPORT_DEVICEUSAGEUSERDETAIL;
    }
    get description() {
        return 'Gets information about Microsoft Teams device usage by user';
    }
    get usageEndpoint() {
        return 'getTeamsDeviceUsageUserDetail';
    }
}
export default new TeamsReportDeviceUsageUserDetailCommand();
//# sourceMappingURL=report-deviceusageuserdetail.js.map