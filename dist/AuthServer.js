import http from 'http';
import { Auth } from './Auth.js';
import { browserUtil } from './utils/browserUtil.js';
import { networkAccess } from './utils/networkAccess.js';
export class AuthServer {
    constructor() {
        this.debug = false;
        this.resource = "";
        this.generatedServerUrl = "";
        this.initializeServer = (connection, resource, resolve, reject, logger, debug = false) => {
            this.connection = connection;
            this.resolve = resolve;
            this.reject = reject;
            this.logger = logger;
            this.debug = debug;
            this.resource = resource;
            this.httpServer = http.createServer(this.httpRequest).listen(0, this.httpListener);
        };
        this.httpListener = async () => {
            const requestState = Math.random().toString(16).substring(2, 22);
            const address = this.httpServer.address();
            this.generatedServerUrl = `http://localhost:${address.port}`;
            const url = `${Auth.getEndpointForResource('https://login.microsoftonline.com', this.connection.cloudType)}/${this.connection.tenant}/oauth2/authorize?response_type=code&client_id=${this.connection.appId}&redirect_uri=${this.generatedServerUrl}&state=${requestState}&resource=${this.resource}&prompt=select_account`;
            if (this.debug) {
                await this.logger.logToStderr('Redirect URL:');
                await this.logger.logToStderr(url);
                await this.logger.logToStderr('');
            }
            await this.openUrl(url);
        };
        this.httpRequest = async (request, response) => {
            if (this.debug) {
                await this.logger.logToStderr('Response:');
                await this.logger.logToStderr(request.url);
                await this.logger.logToStderr('');
            }
            const urlObj = new URL(request.url, 'http://localhost');
            const queryString = Object.fromEntries(urlObj.searchParams.entries());
            const hasCode = queryString.code !== undefined;
            const hasError = queryString.error !== undefined;
            let body = "";
            if (hasCode === true) {
                body = '<p><b>You have logged into CLI for Microsoft 365!</b></p>';
                if (networkAccess.isRestricted()) {
                    body += '<p>You can close this window.</p>';
                }
                else {
                    body += '<script type="text/JavaScript">setTimeout(function(){ window.location = "https://pnp.github.io/cli-microsoft365/"; },10000);</script>';
                    body += '<p>You can close this window, or we will redirect you to the <a href="https://pnp.github.io/cli-microsoft365/">CLI for Microsoft 365</a> documentation in 10 seconds.</p>';
                }
                this.resolve({
                    code: queryString.code,
                    redirectUri: this.generatedServerUrl
                });
            }
            if (hasError === true) {
                const errorMessage = {
                    error: queryString.error,
                    errorDescription: queryString.error_description
                };
                body = "<p>Oops! Microsoft Entra ID replied with an error message.</p>";
                body += `<p>${errorMessage.error}</p>`;
                if (errorMessage.errorDescription !== undefined) {
                    body += `<p>${errorMessage.errorDescription}</p>`;
                }
                this.reject(errorMessage);
            }
            if (hasCode === false && hasError === false) {
                const errorMessage = {
                    error: "invalid request",
                    errorDescription: "An invalid request has been received by the HTTP server"
                };
                body = "<p>Oops! This is an invalid request.</p>";
                body += `<p>${errorMessage.error}</p>`;
                body += `<p>${errorMessage.errorDescription}</p>`;
                this.reject(errorMessage);
            }
            response.writeHead(200, { 'Access-Control-Allow-Origin': '*', 'Content-Type': 'text/html' });
            response.write(`<html><head><title>CLI for Microsoft 365</title></head><body>${body}</body></html>`);
            response.end();
            this.httpServer.close();
        };
    }
    get server() {
        return this.httpServer;
    }
    async openUrl(url) {
        try {
            await browserUtil.open(url);
            await this.logger.logToStderr("To sign in, use the web browser that just has been opened. Please sign-in there.");
        }
        catch {
            const errorResponse = {
                error: "Can't open the default browser",
                errorDescription: "Was not able to open a browser instance. Try again later or use a different authentication method."
            };
            this.reject(errorResponse);
            this.httpServer.close();
        }
    }
}
export default new AuthServer();
//# sourceMappingURL=AuthServer.js.map