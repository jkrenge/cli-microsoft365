import auth from "../Auth.js";
import { CommandError } from "../Command.js";
export const accessToken = {
    isAppOnlyAccessToken(accessToken) {
        let isAppOnlyAccessToken;
        if (!accessToken || accessToken.length === 0) {
            return isAppOnlyAccessToken;
        }
        const chunks = accessToken.split('.');
        if (chunks.length !== 3) {
            return isAppOnlyAccessToken;
        }
        const tokenString = Buffer.from(chunks[1], 'base64').toString();
        try {
            const token = JSON.parse(tokenString);
            isAppOnlyAccessToken = token.idtyp === 'app';
        }
        catch {
            // Do nothing
        }
        return isAppOnlyAccessToken;
    },
    getTenantIdFromAccessToken(accessToken) {
        let tenantId = '';
        if (!accessToken || accessToken.length === 0) {
            return tenantId;
        }
        const chunks = accessToken.split('.');
        if (chunks.length !== 3) {
            return tenantId;
        }
        const tokenString = Buffer.from(chunks[1], 'base64').toString();
        try {
            const token = JSON.parse(tokenString);
            tenantId = token.tid;
        }
        catch {
            // Do nothing
        }
        return tenantId;
    },
    getUserNameFromAccessToken(accessToken) {
        let userName = '';
        if (!accessToken || accessToken.length === 0) {
            return userName;
        }
        const chunks = accessToken.split('.');
        if (chunks.length !== 3) {
            return userName;
        }
        const tokenString = Buffer.from(chunks[1], 'base64').toString();
        try {
            const token = JSON.parse(tokenString);
            // if authenticated using certificate, there is no upn so use
            // app display name instead
            userName = token.upn || token.app_displayname;
        }
        catch {
            // Do nothing
        }
        return userName;
    },
    getUserIdFromAccessToken(accessToken) {
        let userId = '';
        if (!accessToken || accessToken.length === 0) {
            return userId;
        }
        const chunks = accessToken.split('.');
        if (chunks.length !== 3) {
            return userId;
        }
        const tokenString = Buffer.from(chunks[1], 'base64').toString();
        try {
            const token = JSON.parse(tokenString);
            userId = token.oid;
        }
        catch {
            // Do nothing
        }
        return userId;
    },
    getDecodedAccessToken(accessToken) {
        const chunks = accessToken.split('.');
        const headerString = Buffer.from(chunks[0], 'base64').toString();
        const payloadString = Buffer.from(chunks[1], 'base64').toString();
        const header = JSON.parse(headerString);
        const payload = JSON.parse(payloadString);
        return { header, payload };
    },
    getScopesFromAccessToken(accessToken) {
        let scopes = [];
        if (!accessToken || accessToken.length === 0) {
            return scopes;
        }
        const chunks = accessToken.split('.');
        if (chunks.length !== 3) {
            return scopes;
        }
        const tokenString = Buffer.from(chunks[1], 'base64').toString();
        try {
            const token = JSON.parse(tokenString);
            if (token.scp?.length > 0) {
                scopes = token.scp.split(' ');
            }
        }
        catch {
            // Do nothing
        }
        return scopes;
    },
    /**
     * Asserts the presence of a delegated or application-only access token.
     * @throws {CommandError} Will throw an error if the access token is not available.
     * @throws {CommandError} Will throw an error if the access token type is not correct.
     */
    assertAccessTokenType(type) {
        const accessToken = auth?.connection?.accessTokens?.[auth.defaultResource]?.accessToken;
        if (!accessToken) {
            throw new CommandError('No access token found.');
        }
        const isAppAccessToken = this.isAppOnlyAccessToken(accessToken);
        if (type === 'delegated' && isAppAccessToken) {
            throw new CommandError('This command requires delegated permissions.');
        }
        if (type === 'application' && !isAppAccessToken) {
            throw new CommandError('This command requires application-only permissions.');
        }
    }
};
//# sourceMappingURL=accessToken.js.map