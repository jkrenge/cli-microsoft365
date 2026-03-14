import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
class TenantReportSettingsGetCommand extends GraphCommand {
    get name() {
        return commands.REPORT_SETTINGS_GET;
    }
    get description() {
        return 'Get the tenant-level settings for Microsoft 365 reports';
    }
    async commandAction(logger) {
        if (this.verbose) {
            await logger.logToStderr('Getting tenant-level settings for Microsoft 365 reports...');
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/admin/reportSettings`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        try {
            const res = await request.get(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new TenantReportSettingsGetCommand();
//# sourceMappingURL=report-settings-get.js.map