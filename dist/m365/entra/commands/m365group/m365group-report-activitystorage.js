import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
class M365GroupReportActivityStorageCommand extends PeriodBasedReport {
    get name() {
        return commands.M365GROUP_REPORT_ACTIVITYSTORAGE;
    }
    get description() {
        return 'Get the total storage used across all group mailboxes and group sites';
    }
    get usageEndpoint() {
        return 'getOffice365GroupsActivityStorage';
    }
}
export default new M365GroupReportActivityStorageCommand();
//# sourceMappingURL=m365group-report-activitystorage.js.map