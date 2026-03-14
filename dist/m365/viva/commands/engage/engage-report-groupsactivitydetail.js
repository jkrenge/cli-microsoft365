import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';
class VivaEngageReportGroupsActivityDetailCommand extends DateAndPeriodBasedReport {
    get name() {
        return commands.ENGAGE_REPORT_GROUPSACTIVITYDETAIL;
    }
    get usageEndpoint() {
        return 'getYammerGroupsActivityDetail';
    }
    get description() {
        return 'Gets details about Viva Engage groups activity by group';
    }
}
export default new VivaEngageReportGroupsActivityDetailCommand();
//# sourceMappingURL=engage-report-groupsactivitydetail.js.map