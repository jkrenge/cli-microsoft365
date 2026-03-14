import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    displayConcealedNames: z.boolean().alias('d')
});
class TenantReportSettingsSetCommand extends GraphCommand {
    get name() {
        return commands.REPORT_SETTINGS_SET;
    }
    get description() {
        return 'Update tenant-level settings for Microsoft 365 reports';
    }
    get schema() {
        return options;
    }
    async commandAction(logger, args) {
        try {
            const { displayConcealedNames } = args.options;
            if (this.verbose) {
                await logger.logToStderr(`Updating report setting displayConcealedNames to ${displayConcealedNames}`);
            }
            const requestOptions = {
                url: `${this.resource}/v1.0/admin/reportSettings`,
                headers: {
                    accept: 'application/json;odata.metadata=none',
                    'content-type': 'application/json'
                },
                responseType: 'json',
                data: {
                    displayConcealedNames
                }
            };
            await request.patch(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new TenantReportSettingsSetCommand();
//# sourceMappingURL=report-settings-set.js.map