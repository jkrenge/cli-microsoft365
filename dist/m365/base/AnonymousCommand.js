import Command from '../../Command.js';
export default class AnonymousCommand extends Command {
    async action(logger, args) {
        await this.initAction(args, logger);
        await this.commandAction(logger, args);
    }
}
//# sourceMappingURL=AnonymousCommand.js.map