import Command, { CommandError } from "../../Command.js";
import auth from "../../Auth.js";
import { accessToken } from "../../utils/accessToken.js";
export default class VivaEngageCommand extends Command {
    get resource() {
        return 'https://www.yammer.com/api';
    }
    async initAction(args, logger) {
        await super.initAction(args, logger);
        if (!auth.connection.active) {
            // we fail no login in the base command command class
            return;
        }
        accessToken.assertAccessTokenType('delegated');
    }
    handleRejectedODataJsonPromise(response) {
        if (response.statusCode === 404) {
            throw new CommandError("Not found (404)");
        }
        else if (response.error && response.error.base) {
            throw new CommandError(response.error.base);
        }
        else {
            throw new CommandError(response);
        }
    }
}
//# sourceMappingURL=VivaEngageCommand.js.map