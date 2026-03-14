import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class OneDriveReportUsageAccountCountsCommand extends PeriodBasedReport {
    get name() {
        return commands.REPORT_USAGEACCOUNTCOUNTS;
    }
    get usageEndpoint() {
        return 'getOneDriveUsageAccountCounts';
    }
    get description() {
        return 'Gets the trend in the number of active OneDrive for Business sites';
    }
}
export default new OneDriveReportUsageAccountCountsCommand();
//# sourceMappingURL=report-usageaccountcounts.js.map