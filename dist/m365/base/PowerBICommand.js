import Command from '../../Command.js';
import auth from '../../Auth.js';
import { accessToken } from '../../utils/accessToken.js';
export default class PowerBICommand extends Command {
    get resource() {
        return 'https://api.powerbi.com';
    }
    async initAction(args, logger) {
        await super.initAction(args, logger);
        if (!auth.connection.active) {
            // we fail no login in the base command command class
            return;
        }
        accessToken.assertAccessTokenType('delegated');
    }
}
//# sourceMappingURL=PowerBICommand.js.map