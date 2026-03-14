import os from 'os';
import auth from '../../../Auth.js';
import { cli } from '../../../cli/cli.js';
import Command from '../../../Command.js';
import { app } from '../../../utils/app.js';
import { validation } from '../../../utils/validation.js';
import commands from '../commands.js';
class CliDoctorCommand extends Command {
    get name() {
        return commands.DOCTOR;
    }
    get description() {
        return 'Retrieves diagnostic information about the current environment';
    }
    async commandAction(logger) {
        const roles = [];
        const scopes = new Map();
        Object.keys(auth.connection.accessTokens).forEach(resource => {
            const accessToken = auth.connection.accessTokens[resource].accessToken;
            this.getRolesFromAccessToken(accessToken).forEach(role => roles.push(role));
            const [res, scp] = this.getScopesFromAccessToken(accessToken);
            if (res !== "") {
                scopes.set(res, scp);
            }
        });
        const diagnosticInfo = {
            os: {
                platform: os.platform(),
                version: os.version(),
                release: os.release()
            },
            cliVersion: app.packageJson().version,
            nodeVersion: process.version,
            cliEntraAppId: auth.connection.appId,
            cliEntraAppTenant: validation.isValidGuid(auth.connection.tenant) ? 'single' : auth.connection.tenant,
            authMode: auth.connection.authType,
            cliEnvironment: process.env.CLIMICROSOFT365_ENV ? process.env.CLIMICROSOFT365_ENV : '',
            cliConfig: cli.getConfig().all,
            roles: roles,
            scopes: Object.fromEntries(scopes)
        };
        await logger.log(diagnosticInfo);
    }
    getRolesFromAccessToken(accessToken) {
        let roles = [];
        if (!accessToken || accessToken.length === 0) {
            return roles;
        }
        const chunks = accessToken.split('.');
        if (chunks.length !== 3) {
            return roles;
        }
        const tokenString = Buffer.from(chunks[1], 'base64').toString();
        const token = JSON.parse(tokenString);
        if (token.roles !== undefined) {
            roles = token.roles;
        }
        return roles;
    }
    getScopesFromAccessToken(accessToken) {
        let resource = "";
        let scopes = [];
        if (!accessToken || accessToken.length === 0) {
            return [resource, scopes];
        }
        const chunks = accessToken.split('.');
        if (chunks.length !== 3) {
            return [resource, scopes];
        }
        const tokenString = Buffer.from(chunks[1], 'base64').toString();
        const token = JSON.parse(tokenString);
        if (token.scp?.length > 0) {
            resource = token.aud.replace(/(-my|-admin).sharepoint.com/, '.sharepoint.com');
            scopes = token.scp.split(' ');
        }
        return [resource, scopes];
    }
}
export default new CliDoctorCommand();
//# sourceMappingURL=cli-doctor.js.map