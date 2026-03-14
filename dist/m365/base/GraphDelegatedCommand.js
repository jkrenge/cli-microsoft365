import auth from '../../Auth.js';
import { accessToken } from '../../utils/accessToken.js';
import GraphCommand from './GraphCommand.js';
/**
 * This command class is for delegated Graph commands.
 */
export default class GraphDelegatedCommand extends GraphCommand {
    async initAction(args, logger) {
        await super.initAction(args, logger);
        if (!auth.connection.active) {
            // we fail no login in the base command command class
            return;
        }
        accessToken.assertAccessTokenType('delegated');
    }
}
//# sourceMappingURL=GraphDelegatedCommand.js.map