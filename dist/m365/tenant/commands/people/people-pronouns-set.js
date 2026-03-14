import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    enabled: z.boolean().alias('e')
});
class TenantPeoplePronounsSetCommand extends GraphCommand {
    get name() {
        return commands.PEOPLE_PRONOUNS_SET;
    }
    get description() {
        return 'Manage pronouns settings for an organization';
    }
    get schema() {
        return options;
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr('Updating pronouns settings...');
            }
            const requestOptions = {
                url: `${this.resource}/v1.0/admin/people/pronouns`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                data: {
                    isEnabledInOrganization: args.options.enabled
                },
                responseType: 'json'
            };
            const pronouns = await request.patch(requestOptions);
            await logger.log(pronouns);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new TenantPeoplePronounsSetCommand();
//# sourceMappingURL=people-pronouns-set.js.map