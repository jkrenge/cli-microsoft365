import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class TeamsReportUserActivityUserCountsCommand extends PeriodBasedReport {
    get name() {
        return commands.REPORT_USERACTIVITYUSERCOUNTS;
    }
    get description() {
        return 'Get the number of Microsoft Teams users by activity type. The activity types are number of teams chat messages, private chat messages, calls, or meetings.';
    }
    get usageEndpoint() {
        return 'getTeamsUserActivityUserCounts';
    }
}
export default new TeamsReportUserActivityUserCountsCommand();
//# sourceMappingURL=report-useractivityusercounts.js.map