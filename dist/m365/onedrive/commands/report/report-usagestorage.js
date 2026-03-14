import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class OneDriveReportUsageStorageCommand extends PeriodBasedReport {
    get name() {
        return commands.REPORT_USAGESTORAGE;
    }
    get usageEndpoint() {
        return 'getOneDriveUsageStorage';
    }
    get description() {
        return 'Gets the trend on the amount of storage you are using in OneDrive for Business';
    }
}
export default new OneDriveReportUsageStorageCommand();
//# sourceMappingURL=report-usagestorage.js.map