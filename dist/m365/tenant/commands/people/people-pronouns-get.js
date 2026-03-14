import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TenantPeoplePronounsGetCommand extends GraphCommand {
    get name() {
        return commands.PEOPLE_PRONOUNS_GET;
    }
    get description() {
        return 'Retrieves information about pronouns settings for an organization';
    }
    async commandAction(logger) {
        try {
            if (this.verbose) {
                await logger.logToStderr('Retrieving information about pronouns settings...');
            }
            const requestOptions = {
                url: `${this.resource}/v1.0/admin/people/pronouns`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            const pronouns = await request.get(requestOptions);
            await logger.log(pronouns);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new TenantPeoplePronounsGetCommand();
//# sourceMappingURL=people-pronouns-get.js.map