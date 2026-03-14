import { cli } from "./cli/cli.js";
export async function executeCommand(commandName, options, listener) {
    cli.loadAllCommandsInfo();
    await cli.loadCommandFromArgs(commandName.split(' '));
    if (!cli.commandToExecute) {
        throw `Command not found: ${commandName}`;
    }
    return cli.executeCommandWithOutput(cli.commandToExecute.command, { options: options ?? {} }, listener);
}
//# sourceMappingURL=api.js.map