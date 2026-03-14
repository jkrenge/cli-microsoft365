import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';
class VivaEngageReportDeviceUsageUserDetailCommand extends DateAndPeriodBasedReport {
    get name() {
        return commands.ENGAGE_REPORT_DEVICEUSAGEUSERDETAIL;
    }
    get usageEndpoint() {
        return 'getYammerDeviceUsageUserDetail';
    }
    get description() {
        return 'Gets details about Viva Engage device usage by user';
    }
}
export default new VivaEngageReportDeviceUsageUserDetailCommand();
//# sourceMappingURL=engage-report-deviceusageuserdetail.js.map