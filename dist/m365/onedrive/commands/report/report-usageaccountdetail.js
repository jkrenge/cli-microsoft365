import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';
class OneDriveReportUsageAccountDetailCommand extends DateAndPeriodBasedReport {
    get name() {
        return commands.REPORT_USAGEACCOUNTDETAIL;
    }
    get usageEndpoint() {
        return 'getOneDriveUsageAccountDetail';
    }
    get description() {
        return 'Gets details about OneDrive usage by account';
    }
}
export default new OneDriveReportUsageAccountDetailCommand();
//# sourceMappingURL=report-usageaccountdetail.js.map