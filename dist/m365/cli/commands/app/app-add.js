import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import config from '../../../../config.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';
import { entraApp } from '../../../../utils/entraApp.js';
import { accessToken } from '../../../../utils/accessToken.js';
import auth from '../../../../Auth.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    name: z.string().optional().default('CLI for M365').alias('n'),
    scopes: z.string().optional().default('minimal').alias('s'),
    saveToConfig: z.boolean().optional()
});
class CliAppAddCommand extends GraphCommand {
    get name() {
        return commands.APP_ADD;
    }
    get description() {
        return 'Creates a Microsoft Entra application registration for CLI for Microsoft 365';
    }
    get schema() {
        return options;
    }
    getRefinedSchema(schema) {
        return schema
            .refine(options => {
            const scopes = options.scopes;
            if (!scopes.includes(',')) {
                return scopes === 'minimal' || scopes === 'all';
            }
            const scopeList = scopes.split(',').map(s => s.trim());
            return scopeList.every(scope => scope.startsWith('https'));
        }, {
            message: "Scopes must be 'minimal', 'all', or comma-separated list of URLs starting with 'https'. 'minimal' and 'all' cannot be combined with other scopes.",
            path: ['scopes']
        });
    }
    async commandAction(logger, args) {
        try {
            const options = {
                allowPublicClientFlows: true,
                apisDelegated: this.getScopes(args.options),
                implicitFlow: false,
                multitenant: false,
                name: args.options.name,
                platform: 'publicClient',
                redirectUris: 'http://localhost,https://localhost,https://login.microsoftonline.com/common/oauth2/nativeclient'
            };
            const apis = await entraApp.resolveApis({
                options,
                logger,
                verbose: this.verbose,
                debug: this.debug
            });
            const appInfo = await entraApp.createAppRegistration({
                options,
                unknownOptions: {},
                apis,
                logger,
                verbose: this.verbose,
                debug: this.debug
            });
            appInfo.tenantId = accessToken.getTenantIdFromAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken);
            await entraApp.grantAdminConsent({
                appInfo,
                appPermissions: entraApp.appPermissions,
                adminConsent: true,
                logger,
                debug: this.debug
            });
            if (args.options.saveToConfig) {
                cli.getConfig().set('clientId', appInfo.appId);
                cli.getConfig().set('tenantId', appInfo.tenantId);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getScopes(options) {
        if (options.scopes === 'all') {
            return config.allScopes.join(',');
        }
        else if (options.scopes === 'minimal') {
            return config.minimalScopes.join(',');
        }
        return options.scopes;
    }
}
export default new CliAppAddCommand();
//# sourceMappingURL=app-add.js.map