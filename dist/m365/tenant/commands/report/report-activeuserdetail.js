import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';
class TenantReportActiveUserDetailCommand extends DateAndPeriodBasedReport {
    get name() {
        return commands.REPORT_ACTIVEUSERDETAIL;
    }
    get description() {
        return 'Gets details about Microsoft 365 active users';
    }
    get usageEndpoint() {
        return 'getOffice365ActiveUserDetail';
    }
}
export default new TenantReportActiveUserDetailCommand();
//# sourceMappingURL=report-activeuserdetail.js.map