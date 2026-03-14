import { cli } from "../../../../cli/cli.js";
import AnonymousCommand from "../../../base/AnonymousCommand.js";
import commands from "../../commands.js";
class CliConfigListCommand extends AnonymousCommand {
    get name() {
        return commands.CONFIG_LIST;
    }
    get description() {
        return 'List all self set CLI for Microsoft 365 configurations';
    }
    async commandAction(logger) {
        await logger.log(cli.getConfig().all);
    }
}
export default new CliConfigListCommand();
//# sourceMappingURL=config-list.js.map