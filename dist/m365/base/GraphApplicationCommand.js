import auth from '../../Auth.js';
import { accessToken } from '../../utils/accessToken.js';
import GraphCommand from './GraphCommand.js';
/**
 * This command class is for application-only Graph commands.
 */
export default class GraphApplicationCommand extends GraphCommand {
    async initAction(args, logger) {
        await super.initAction(args, logger);
        if (!auth.connection.active) {
            // we fail no login in the base command command class
            return;
        }
        accessToken.assertAccessTokenType('application');
    }
}
//# sourceMappingURL=GraphApplicationCommand.js.map