import commands from '../../commands.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import { odata } from '../../../../utils/odata.js';
class SpeContainerTypeListCommand extends GraphDelegatedCommand {
    get name() {
        return commands.CONTAINERTYPE_LIST;
    }
    get description() {
        return 'Lists all container types';
    }
    defaultProperties() {
        return ['id', 'name', 'owningAppId'];
    }
    async commandAction(logger) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Retrieving list of Container types...`);
            }
            const containerTypes = await odata.getAllItems(`${this.resource}/beta/storage/fileStorage/containerTypes`);
            await logger.log(containerTypes);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new SpeContainerTypeListCommand();
//# sourceMappingURL=containertype-list.js.map