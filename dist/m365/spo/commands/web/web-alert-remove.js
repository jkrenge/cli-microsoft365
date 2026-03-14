import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
const options = z.strictObject({
    ...globalOptionsZod.shape,
    webUrl: z.string().alias('u')
        .refine(url => validation.isValidSharePointUrl(url) === true, {
        error: e => `'${e.input}' is not a valid SharePoint URL.`
    }),
    id: z.uuid(),
    force: z.boolean().optional().alias('f')
});
class SpoWebAlertRemoveCommand extends SpoCommand {
    get name() {
        return commands.WEB_ALERT_REMOVE;
    }
    get description() {
        return 'Removes an alert from a SharePoint list';
    }
    get schema() {
        return options;
    }
    async commandAction(logger, args) {
        if (!args.options.force) {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the alert with id '${args.options.id}' from site '${args.options.webUrl}'?` });
            if (!result) {
                return;
            }
        }
        try {
            if (this.verbose) {
                await logger.logToStderr(`Removing alert with ID '${args.options.id}' from site '${args.options.webUrl}'...`);
            }
            const requestOptions = {
                url: `${args.options.webUrl}/_api/web/Alerts/DeleteAlert('${formatting.encodeQueryParameter(args.options.id)}')`,
                headers: {
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            await request.delete(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new SpoWebAlertRemoveCommand();
//# sourceMappingURL=web-alert-remove.js.map