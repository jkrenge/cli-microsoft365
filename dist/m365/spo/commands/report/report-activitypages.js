import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class SpoReportActivityPagesCommand extends PeriodBasedReport {
    get name() {
        return commands.REPORT_ACTIVITYPAGES;
    }
    get usageEndpoint() {
        return 'getSharePointActivityPages';
    }
    get description() {
        return 'Gets the number of unique pages visited by users';
    }
}
export default new SpoReportActivityPagesCommand();
//# sourceMappingURL=report-activitypages.js.map