import fs from 'fs';
import { z } from 'zod';
import { cli } from '../../cli/cli.js';
import Command, { CommandError, globalOptionsZod } from '../../Command.js';
import { formatting } from '../../utils/formatting.js';
export const appCommandOptions = z.object({
    ...globalOptionsZod.shape,
    appId: z.uuid().optional()
});
export default class AppCommand extends Command {
    get resource() {
        return 'https://graph.microsoft.com';
    }
    get schema() {
        return appCommandOptions;
    }
    async action(logger, args) {
        const m365rcJsonPath = '.m365rc.json';
        if (!fs.existsSync(m365rcJsonPath)) {
            throw new CommandError(`Could not find file: ${m365rcJsonPath}`);
        }
        try {
            const m365rcJsonContents = fs.readFileSync(m365rcJsonPath, 'utf8');
            if (!m365rcJsonContents) {
                throw new CommandError(`File ${m365rcJsonPath} is empty`);
            }
            this.m365rcJson = JSON.parse(m365rcJsonContents);
        }
        catch (err) {
            if (err instanceof CommandError) {
                throw err;
            }
            throw new CommandError(`Could not parse file: ${m365rcJsonPath}`);
        }
        if (!this.m365rcJson.apps ||
            this.m365rcJson.apps.length === 0) {
            throw new CommandError(`No Entra apps found in ${m365rcJsonPath}`);
        }
        if (args.options.appId) {
            if (!this.m365rcJson.apps.some(app => app.appId === args.options.appId)) {
                throw new CommandError(`App ${args.options.appId} not found in ${m365rcJsonPath}`);
            }
            this.appId = args.options.appId;
            return super.action(logger, args);
        }
        if (this.m365rcJson.apps.length === 1) {
            this.appId = this.m365rcJson.apps[0].appId;
            return super.action(logger, args);
        }
        if (this.m365rcJson.apps.length > 1) {
            this.m365rcJson.apps.forEach((app, index) => {
                app.appIdIndex = index;
            });
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('appId', this.m365rcJson.apps);
            const result = await cli.handleMultipleResultsFound(`Multiple Entra apps found in ${m365rcJsonPath}.`, resultAsKeyValuePair);
            this.appId = result.appId;
            await super.action(logger, args);
        }
    }
}
//# sourceMappingURL=AppCommand.js.map