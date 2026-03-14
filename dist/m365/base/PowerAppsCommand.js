import auth, { CloudType } from '../../Auth.js';
import Command, { CommandError } from '../../Command.js';
import { accessToken } from '../../utils/accessToken.js';
export default class PowerAppsCommand extends Command {
    get resource() {
        return 'https://api.powerapps.com';
    }
    async initAction(args, logger) {
        await super.initAction(args, logger);
        if (!auth.connection.active) {
            // we fail no login in the base command command class
            return;
        }
        if (auth.connection.cloudType !== CloudType.Public) {
            throw new CommandError(`Power Apps commands only support the public cloud at the moment. We'll add support for other clouds in the future. Sorry for the inconvenience.`);
        }
        accessToken.assertAccessTokenType('delegated');
    }
}
//# sourceMappingURL=PowerAppsCommand.js.map