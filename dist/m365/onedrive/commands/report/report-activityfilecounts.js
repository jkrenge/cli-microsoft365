import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class OneDriveReportActivityFileCountCommand extends PeriodBasedReport {
    get name() {
        return commands.REPORT_ACTIVITYFILECOUNTS;
    }
    get usageEndpoint() {
        return 'getOneDriveActivityFileCounts';
    }
    get description() {
        return 'Gets the number of unique, licensed users that performed file interactions against any OneDrive account';
    }
}
export default new OneDriveReportActivityFileCountCommand();
//# sourceMappingURL=report-activityfilecounts.js.map