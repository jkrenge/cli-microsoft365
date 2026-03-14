import fs from 'fs';
import path from 'path';
import { cli } from '../../../../cli/cli.js';
import { CommandError } from '../../../../Command.js';
import spoServicePrincipalGrantAddCommand from '../../../spo/commands/serviceprincipal/serviceprincipal-grant-add.js';
import commands from '../../commands.js';
import { BaseProjectCommand } from './base-project-command.js';
class SpfxProjectPermissionSGrantCommand extends BaseProjectCommand {
    get name() {
        return commands.PROJECT_PERMISSIONS_GRANT;
    }
    get description() {
        return 'Grant API permissions defined in the current SPFx project';
    }
    constructor() {
        super();
    }
    async commandAction(logger) {
        this.projectRootPath = this.getProjectRoot(process.cwd());
        if (this.projectRootPath === null) {
            throw new CommandError(`Couldn't find project root folder`, SpfxProjectPermissionSGrantCommand.ERROR_NO_PROJECT_ROOT_FOLDER);
        }
        if (this.debug) {
            await logger.logToStderr(`Granting API permissions defined in the current SPFx project`);
        }
        try {
            const webApiPermissionsRequest = this.getWebApiPermissionRequest(path.join(this.projectRootPath, 'config', 'package-solution.json'));
            for (const permission of webApiPermissionsRequest) {
                const options = {
                    resource: permission.resource,
                    scope: permission.scope,
                    output: 'json',
                    debug: this.debug,
                    verbose: this.verbose
                };
                let output = null;
                try {
                    output = await cli.executeCommandWithOutput(spoServicePrincipalGrantAddCommand, { options: { ...options, _: [] } });
                }
                catch (err) {
                    if (err.error && err.error.message.indexOf('already exists') > -1) {
                        await this.warn(logger, err.error.message);
                        continue;
                    }
                    else {
                        throw err;
                    }
                }
                const getGrantOutput = JSON.parse(output.stdout);
                await logger.log(getGrantOutput);
            }
        }
        catch (error) {
            throw new CommandError(error);
        }
    }
    getWebApiPermissionRequest(filePath) {
        if (!fs.existsSync(filePath)) {
            throw (`The package-solution.json file could not be found`);
        }
        const existingContent = fs.readFileSync(filePath, 'utf-8');
        const solutionContent = JSON.parse(existingContent);
        return solutionContent.solution.webApiPermissionRequests;
    }
}
SpfxProjectPermissionSGrantCommand.ERROR_NO_PROJECT_ROOT_FOLDER = 1;
export default new SpfxProjectPermissionSGrantCommand();
//# sourceMappingURL=project-permissions-grant.js.map