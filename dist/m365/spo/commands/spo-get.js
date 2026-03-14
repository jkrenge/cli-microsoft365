import auth from '../../../Auth.js';
import SpoCommand from '../../base/SpoCommand.js';
import commands from '../commands.js';
class SpoGetCommand extends SpoCommand {
    get name() {
        return commands.GET;
    }
    get description() {
        return 'Gets the context URL for the root SharePoint site collection and SharePoint tenant admin site';
    }
    async commandAction(logger) {
        const spoContext = {
            SpoUrl: auth.connection.spoUrl ? auth.connection.spoUrl : ''
        };
        await logger.log(spoContext);
    }
}
export default new SpoGetCommand();
//# sourceMappingURL=spo-get.js.map