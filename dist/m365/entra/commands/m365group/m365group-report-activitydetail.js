import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';
class M365GroupReportActivityDetailCommand extends DateAndPeriodBasedReport {
    get name() {
        return commands.M365GROUP_REPORT_ACTIVITYDETAIL;
    }
    get description() {
        return 'Get details about Microsoft 365 Groups activity by group';
    }
    get usageEndpoint() {
        return 'getOffice365GroupsActivityDetail';
    }
}
export default new M365GroupReportActivityDetailCommand();
//# sourceMappingURL=m365group-report-activitydetail.js.map