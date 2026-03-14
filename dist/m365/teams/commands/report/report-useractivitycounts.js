import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class TeamsReportUserActivityCountsCommand extends PeriodBasedReport {
    get name() {
        return commands.REPORT_USERACTIVITYCOUNTS;
    }
    get description() {
        return 'Get the number of Microsoft Teams activities by activity type. The activity types are team chat messages, private chat messages, calls, and meetings.';
    }
    get usageEndpoint() {
        return 'getTeamsUserActivityCounts';
    }
}
export default new TeamsReportUserActivityCountsCommand();
//# sourceMappingURL=report-useractivitycounts.js.map