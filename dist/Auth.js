import { AzureCloudInstance } from '@azure/msal-common';
import assert from 'assert';
import { CommandError } from './Command.js';
import { FileTokenStorage } from './auth/FileTokenStorage.js';
import { MsalNetworkClient } from './auth/MsalNetworkClient.js';
import { msalCachePlugin } from './auth/msalCachePlugin.js';
import { cli } from './cli/cli.js';
import request from './request.js';
import { settingsNames } from './settingsNames.js';
import * as accessTokenUtil from './utils/accessToken.js';
import { browserUtil } from './utils/browserUtil.js';
export var CloudType;
(function (CloudType) {
    CloudType["Public"] = "Public";
    CloudType["USGov"] = "USGov";
    CloudType["USGovHigh"] = "USGovHigh";
    CloudType["USGovDoD"] = "USGovDoD";
    CloudType["China"] = "China";
})(CloudType || (CloudType = {}));
export class Connection {
    constructor() {
        this.active = false;
        this.authType = AuthType.DeviceCode;
        this.certificateType = CertificateType.Unknown;
        // ID of the tenant where the Microsoft Entra app is registered; common if multi-tenant
        this.tenant = 'common';
        this.cloudType = CloudType.Public;
        this.accessTokens = {};
        this.cloudType = CloudType.Public;
    }
    deactivate() {
        this.active = false;
        this.name = undefined;
        this.identityName = undefined;
        this.identityId = undefined;
        this.identityTenantId = undefined;
        this.accessTokens = {};
        this.authType = AuthType.DeviceCode;
        this.userName = undefined;
        this.password = undefined;
        this.certificateType = CertificateType.Unknown;
        this.cloudType = CloudType.Public;
        this.certificate = undefined;
        this.thumbprint = undefined;
        this.spoUrl = undefined;
        this.spoTenantId = undefined;
        this.appId = cli.getClientId();
        this.tenant = cli.getTenant();
    }
}
export var AuthType;
(function (AuthType) {
    AuthType["DeviceCode"] = "deviceCode";
    AuthType["Password"] = "password";
    AuthType["Certificate"] = "certificate";
    AuthType["Identity"] = "identity";
    AuthType["FederatedIdentity"] = "federatedIdentity";
    AuthType["Browser"] = "browser";
    AuthType["Secret"] = "secret";
})(AuthType || (AuthType = {}));
export var CertificateType;
(function (CertificateType) {
    CertificateType[CertificateType["Unknown"] = 0] = "Unknown";
    CertificateType[CertificateType["Base64"] = 1] = "Base64";
    CertificateType[CertificateType["Binary"] = 2] = "Binary";
})(CertificateType || (CertificateType = {}));
export class Auth {
    // Retrieves the connections from the file store if it's not already loaded
    async getAllConnections() {
        if (this._allConnections === undefined) {
            try {
                this._allConnections = await this.getAllConnectionsFromStorage();
            }
            catch {
                this._allConnections = [];
            }
        }
        return this._allConnections;
    }
    get connection() {
        return this._connection;
    }
    get defaultResource() {
        return Auth.getEndpointForResource('https://graph.microsoft.com', this._connection.cloudType);
    }
    constructor() {
        this._connection = new Connection();
    }
    // we need to init cloud endpoints here, because we're using CloudType enum
    // as indexers, which we can't do in the static initializer
    // it also needs to be a separate method that we call here, because in tests
    // we're mocking auth and calling its constructor
    static initialize() {
        this.cloudEndpoints[CloudType.USGov] = {
            'https://graph.microsoft.com': 'https://graph.microsoft.com',
            'https://management.azure.com/': 'https://management.usgovcloudapi.net/',
            'https://login.microsoftonline.com': 'https://login.microsoftonline.com'
        };
        this.cloudEndpoints[CloudType.USGovHigh] = {
            'https://graph.microsoft.com': 'https://graph.microsoft.us',
            'https://management.azure.com/': 'https://management.usgovcloudapi.net/',
            'https://login.microsoftonline.com': 'https://login.microsoftonline.us'
        };
        this.cloudEndpoints[CloudType.USGovDoD] = {
            'https://graph.microsoft.com': 'https://dod-graph.microsoft.us',
            'https://management.azure.com/': 'https://management.usgovcloudapi.net/',
            'https://login.microsoftonline.com': 'https://login.microsoftonline.us'
        };
        this.cloudEndpoints[CloudType.China] = {
            'https://graph.microsoft.com': 'https://microsoftgraph.chinacloudapi.cn',
            'https://management.azure.com/': 'https://management.chinacloudapi.cn',
            'https://login.microsoftonline.com': 'https://login.chinacloudapi.cn'
        };
    }
    async restoreAuth() {
        // check if auth has been restored previously
        if (this._connection.active) {
            return;
        }
        try {
            const connection = await this.getConnectionInfoFromStorage();
            this._connection = Object.assign(this._connection, connection);
        }
        catch {
            // Do nothing
        }
    }
    async ensureAccessToken(resource, logger, debug = false, fetchNew = false) {
        const now = new Date();
        const accessToken = this.connection.accessTokens[resource];
        const expiresOn = accessToken && accessToken.expiresOn ?
            // if expiresOn is serialized from the service file, it's set as a string
            // if it's coming from MSAL, it's a Date
            typeof accessToken.expiresOn === 'string' ? new Date(accessToken.expiresOn) : accessToken.expiresOn
            : new Date(0);
        if (!fetchNew && accessToken && expiresOn > now) {
            if (debug) {
                await logger.logToStderr(`Existing access token ${accessToken.accessToken} still valid. Returning...`);
            }
            return accessToken.accessToken;
        }
        else {
            if (debug) {
                if (!accessToken) {
                    await logger.logToStderr(`No token found for resource ${resource}.`);
                }
                else {
                    await logger.logToStderr(`Access token expired. Token: ${accessToken.accessToken}, ExpiresAt: ${accessToken.expiresOn}`);
                }
            }
        }
        let getTokenPromise;
        // When using an application identity, you can't retrieve the access token silently, because there is
        // no account. Also (for cert auth) clientApplication is instantiated later
        // after inspecting the specified cert and calculating thumbprint if one
        // wasn't specified
        if (this.connection.authType !== AuthType.Certificate &&
            this.connection.authType !== AuthType.Secret &&
            this.connection.authType !== AuthType.Identity &&
            this.connection.authType !== AuthType.FederatedIdentity) {
            this.clientApplication = await this.getPublicClient(logger, debug);
            if (this.clientApplication) {
                const accounts = await this.clientApplication.getTokenCache().getAllAccounts();
                // if there is an account in the cache and it's active, we can try to get the token silently
                if (accounts.filter(a => a.localAccountId === this.connection.identityId).length > 0 && this.connection.active === true) {
                    getTokenPromise = this.ensureAccessTokenSilent.bind(this);
                }
            }
        }
        if (!getTokenPromise) {
            switch (this.connection.authType) {
                case AuthType.DeviceCode:
                    getTokenPromise = this.ensureAccessTokenWithDeviceCode.bind(this);
                    break;
                case AuthType.Password:
                    getTokenPromise = this.ensureAccessTokenWithPassword.bind(this);
                    break;
                case AuthType.Certificate:
                    getTokenPromise = this.ensureAccessTokenWithCertificate.bind(this);
                    break;
                case AuthType.Identity:
                    getTokenPromise = this.ensureAccessTokenWithIdentity.bind(this);
                    break;
                case AuthType.FederatedIdentity:
                    getTokenPromise = this.ensureAccessTokenWithFederatedIdentity.bind(this);
                    break;
                case AuthType.Browser:
                    getTokenPromise = this.ensureAccessTokenWithBrowser.bind(this);
                    break;
                case AuthType.Secret:
                    getTokenPromise = this.ensureAccessTokenWithSecret.bind(this);
                    break;
            }
        }
        const response = await getTokenPromise(resource, logger, debug, fetchNew);
        if (!response) {
            if (debug) {
                await logger.logToStderr('getTokenPromise authentication result is null.');
            }
            throw 'Failed to retrieve an access token. Please try again.';
        }
        else {
            if (debug) {
                await logger.logToStderr('Response');
                await logger.logToStderr(response);
                await logger.logToStderr('');
            }
        }
        this.connection.accessTokens[resource] = {
            expiresOn: response.expiresOn,
            accessToken: response.accessToken
        };
        this.connection.active = true;
        this.connection.identityName = accessTokenUtil.accessToken.getUserNameFromAccessToken(response.accessToken);
        this.connection.identityId = accessTokenUtil.accessToken.getUserIdFromAccessToken(response.accessToken);
        this.connection.identityTenantId = accessTokenUtil.accessToken.getTenantIdFromAccessToken(response.accessToken);
        this.connection.name = this.connection.name || this.connection.identityId;
        try {
            await this.storeConnectionInfo();
        }
        catch (ex) {
            // error could happen due to an issue with persisting the access
            // token which shouldn't fail the overall token retrieval process
            if (debug) {
                await logger.logToStderr(new CommandError(ex));
            }
        }
        return response.accessToken;
    }
    async getAuthClientConfiguration(logger, debug, certificateThumbprint, certificatePrivateKey, clientSecret) {
        const msal = await import('@azure/msal-node');
        const { LogLevel } = msal;
        const cert = !certificateThumbprint ? undefined : {
            thumbprint: certificateThumbprint,
            privateKey: certificatePrivateKey
        };
        let azureCloudInstance = AzureCloudInstance.None;
        switch (this.connection.cloudType) {
            case CloudType.Public:
            case CloudType.USGov:
                azureCloudInstance = AzureCloudInstance.AzurePublic;
                break;
            case CloudType.China:
                azureCloudInstance = AzureCloudInstance.AzureChina;
                break;
            case CloudType.USGovHigh:
            case CloudType.USGovDoD:
                azureCloudInstance = AzureCloudInstance.AzureUsGovernment;
                break;
        }
        const config = {
            clientId: this.connection.appId,
            authority: `${Auth.getEndpointForResource('https://login.microsoftonline.com', this.connection.cloudType)}/${this.connection.tenant}`,
            azureCloudOptions: {
                azureCloudInstance,
                tenant: this.connection.tenant
            }
        };
        const authConfig = cert
            ? { ...config, clientCertificate: cert }
            : { ...config, clientSecret };
        return {
            auth: authConfig,
            cache: {
                cachePlugin: msalCachePlugin
            },
            system: {
                loggerOptions: {
                    // loggerCallback is called by MSAL which we're not testing
                    /* c8 ignore next 4 */
                    loggerCallback: async (level, message) => {
                        if (level === LogLevel.Error || debug) {
                            await logger.logToStderr(message);
                        }
                    },
                    piiLoggingEnabled: false,
                    logLevel: debug ? LogLevel.Verbose : LogLevel.Error
                },
                networkClient: new MsalNetworkClient()
            }
        };
    }
    async getPublicClient(logger, debug) {
        const msal = await import('@azure/msal-node');
        const { PublicClientApplication } = msal;
        if (this.connection.authType === AuthType.Password &&
            this.connection.tenant === 'common') {
            // common is not supported for the password flow and must be changed to
            // organizations
            this.connection.tenant = 'organizations';
        }
        return new PublicClientApplication(await this.getAuthClientConfiguration(logger, debug));
    }
    async getConfidentialClient(logger, debug, certificateThumbprint, certificatePrivateKey, clientSecret) {
        const msal = await import('@azure/msal-node');
        const { ConfidentialClientApplication } = msal;
        return new ConfidentialClientApplication(await this.getAuthClientConfiguration(logger, debug, certificateThumbprint, certificatePrivateKey, clientSecret));
    }
    retrieveAuthCodeWithBrowser(resource, logger, debug) {
        return new Promise((resolve, reject) => {
            // _authServer is never set before hitting this line, but this check
            // is implemented so that we can support lazy loading
            // but also stub it for testing
            (async () => {
                /* c8 ignore next 3 */
                if (!this._authServer) {
                    this._authServer = (await import('./AuthServer.js')).default;
                }
                this._authServer.initializeServer(this.connection, resource, resolve, reject, logger, debug);
            })().catch(reject);
        });
    }
    async ensureAccessTokenWithBrowser(resource, logger, debug) {
        if (debug) {
            await logger.logToStderr(`Retrieving new access token using interactive browser session...`);
        }
        const response = await this.retrieveAuthCodeWithBrowser(resource, logger, debug);
        if (debug) {
            await logger.logToStderr(`The service returned the code '${response.code}'`);
        }
        return this.clientApplication.acquireTokenByCode({
            code: response.code,
            redirectUri: response.redirectUri,
            scopes: [`${resource}/.default`]
        });
    }
    async ensureAccessTokenSilent(resource, logger, debug, fetchNew) {
        if (debug) {
            await logger.logToStderr(`Retrieving new access token silently`);
        }
        // Asserting identityId because it is expected to be available at this point.
        assert(this.connection.identityId !== undefined);
        const account = await this.clientApplication
            .getTokenCache().getAccountByLocalId(this.connection.identityId);
        // Asserting account because it is expected to be available at this point.
        assert(account !== null);
        return this.clientApplication.acquireTokenSilent({
            account: account,
            scopes: [`${resource}/.default`],
            forceRefresh: fetchNew
        });
    }
    async ensureAccessTokenWithDeviceCode(resource, logger, debug) {
        if (debug) {
            await logger.logToStderr(`Starting Auth.ensureAccessTokenWithDeviceCode. resource: ${resource}, debug: ${debug}`);
        }
        this.deviceCodeRequest = {
            // deviceCodeCallback is called by MSAL which we're not testing
            /* c8 ignore next 1 */
            deviceCodeCallback: response => this.processDeviceCodeCallback(response, logger, debug),
            scopes: [`${resource}/.default`]
        };
        return this.clientApplication.acquireTokenByDeviceCode(this.deviceCodeRequest);
    }
    async processDeviceCodeCallback(response, logger, debug) {
        if (debug) {
            await logger.logToStderr('Response:');
            await logger.logToStderr(response);
            await logger.logToStderr('');
        }
        if (response.message) {
            await logger.logToStderr(`🌶️  ${response.message}`);
        }
        if (cli.getSettingWithDefaultValue(settingsNames.autoOpenLinksInBrowser, false)) {
            await browserUtil.open(response.verificationUri);
        }
        if (cli.getSettingWithDefaultValue(settingsNames.copyDeviceCodeToClipboard, false)) {
            // _clipboardy is never set before hitting this line, but this check
            // is implemented so that we can support lazy loading
            // but also stub it for testing
            /* c8 ignore next 3 */
            if (!this._clipboardy) {
                this._clipboardy = (await import('clipboardy')).default;
            }
            this._clipboardy.writeSync(response.userCode);
        }
    }
    async ensureAccessTokenWithPassword(resource, logger, debug) {
        if (debug) {
            await logger.logToStderr(`Retrieving new access token using credentials...`);
        }
        return this.clientApplication.acquireTokenByUsernamePassword({
            username: this.connection.userName,
            password: this.connection.password,
            scopes: [`${resource}/.default`]
        });
    }
    async ensureAccessTokenWithCertificate(resource, logger, debug, fetchNew) {
        const nodeForge = (await import('node-forge')).default;
        const { pem, pki, asn1, pkcs12 } = nodeForge;
        if (debug) {
            await logger.logToStderr(`Retrieving new access token using certificate...`);
        }
        let cert = '';
        const buf = Buffer.from(this.connection.certificate, 'base64');
        if (this.connection.certificateType === CertificateType.Unknown || this.connection.certificateType === CertificateType.Base64) {
            // First time this method is called, we don't know if certificate is PEM or PFX (type is Unknown)
            // We assume it is PEM but when parsing of PEM fails, we assume it could be PFX
            // Type is persisted on service so subsequent calls only run through the correct parsing flow
            try {
                cert = buf.toString('utf8');
                const pemObjs = pem.decode(cert);
                if (this.connection.thumbprint === undefined) {
                    const pemCertObj = pemObjs.find(pem => pem.type === "CERTIFICATE");
                    const pemCertStr = pem.encode(pemCertObj);
                    const pemCert = pki.certificateFromPem(pemCertStr);
                    this.connection.thumbprint = await this.calculateThumbprint(pemCert);
                }
            }
            catch {
                this.connection.certificateType = CertificateType.Binary;
            }
        }
        if (this.connection.certificateType === CertificateType.Binary) {
            const p12Asn1 = asn1.fromDer(buf.toString('binary'), false);
            const p12Parsed = pkcs12.pkcs12FromAsn1(p12Asn1, false, this.connection.password);
            let keyBags = p12Parsed.getBags({ bagType: pki.oids.pkcs8ShroudedKeyBag });
            const pkcs8ShroudedKeyBag = keyBags[pki.oids.pkcs8ShroudedKeyBag][0];
            if (debug) {
                // check if there is something in the keyBag as well as
                // the pkcs8ShroudedKeyBag. This will give us more information
                // whether there is a cert that can potentially store keys in the keyBag.
                // I could not find a way to add something to the keyBag with all 
                // my attempts, but lets keep it here for troubleshooting purposes.
                await logger.logToStderr(`pkcs8ShroudedKeyBagkeyBags length is ${[pki.oids.pkcs8ShroudedKeyBag].length}`);
                keyBags = p12Parsed.getBags({ bagType: pki.oids.keyBag });
                await logger.logToStderr(`keyBag length is ${keyBags[pki.oids.keyBag].length}`);
            }
            // convert a Forge private key to an ASN.1 RSAPrivateKey
            const rsaPrivateKey = pki.privateKeyToAsn1(pkcs8ShroudedKeyBag.key);
            // wrap an RSAPrivateKey ASN.1 object in a PKCS#8 ASN.1 PrivateKeyInfo
            const privateKeyInfo = pki.wrapRsaPrivateKey(rsaPrivateKey);
            // convert a PKCS#8 ASN.1 PrivateKeyInfo to PEM
            cert = pki.privateKeyInfoToPem(privateKeyInfo);
            if (this.connection.thumbprint === undefined) {
                const certBags = p12Parsed.getBags({ bagType: pki.oids.certBag });
                const certBag = (certBags[pki.oids.certBag])[0];
                this.connection.thumbprint = await this.calculateThumbprint(certBag.cert);
            }
        }
        this.clientApplication = await this.getConfidentialClient(logger, debug, this.connection.thumbprint, cert);
        return this.clientApplication.acquireTokenByClientCredential({
            scopes: [`${resource}/.default`],
            skipCache: fetchNew
        });
    }
    async ensureAccessTokenWithIdentity(resource, logger, debug) {
        const userName = this.connection.userName;
        if (debug) {
            await logger.logToStderr('Will try to retrieve access token using identity...');
        }
        const requestOptions = {
            url: '',
            headers: {
                accept: 'application/json',
                Metadata: true,
                'x-anonymous': true
            },
            responseType: 'json'
        };
        if (process.env.IDENTITY_ENDPOINT && process.env.IDENTITY_HEADER) {
            if (debug) {
                await logger.logToStderr('IDENTITY_ENDPOINT and IDENTITY_HEADER env variables found it is Azure Function, WebApp...');
            }
            requestOptions.url = `${process.env.IDENTITY_ENDPOINT}?resource=${encodeURIComponent(resource)}&api-version=2019-08-01`;
            requestOptions.headers['X-IDENTITY-HEADER'] = process.env.IDENTITY_HEADER;
        }
        else if (process.env.MSI_ENDPOINT && process.env.MSI_SECRET) {
            if (debug) {
                await logger.logToStderr('MSI_ENDPOINT and MSI_SECRET env variables found it is Azure Function or WebApp, but using the old names of the env variables...');
            }
            requestOptions.url = `${process.env.MSI_ENDPOINT}?resource=${encodeURIComponent(resource)}&api-version=2019-08-01`;
            requestOptions.headers['X-IDENTITY-HEADER'] = process.env.MSI_SECRET;
        }
        else if (process.env.IDENTITY_ENDPOINT) {
            if (debug) {
                await logger.logToStderr('IDENTITY_ENDPOINT env variable found it is Azure Could Shell...');
            }
            if (userName && process.env.ACC_CLOUD) {
                // reject for now since the Azure Cloud Shell does not support user-managed identity 
                throw 'Azure Cloud Shell does not support user-managed identity. You can execute the command without the --userName option to login with user identity';
            }
            requestOptions.url = `${process.env.IDENTITY_ENDPOINT}?resource=${encodeURIComponent(resource)}`;
        }
        else if (process.env.MSI_ENDPOINT) {
            if (debug) {
                await logger.logToStderr('MSI_ENDPOINT env variable found it is Azure Could Shell, but using the old names of the env variables...');
            }
            if (userName && process.env.ACC_CLOUD) {
                // reject for now since the Azure Cloud Shell does not support user-managed identity 
                throw 'Azure Cloud Shell does not support user-managed identity. You can execute the command without the --userName option to login with user identity';
            }
            requestOptions.url = `${process.env.MSI_ENDPOINT}?resource=${encodeURIComponent(resource)}`;
        }
        else {
            if (debug) {
                await logger.logToStderr('IDENTITY_ENDPOINT and MSI_ENDPOINT env variables not found. Attempt to get Managed Identity token by using the Azure Virtual Machine API...');
            }
            requestOptions.url = `http://169.254.169.254/metadata/identity/oauth2/token?resource=${encodeURIComponent(resource)}&api-version=2018-02-01`;
        }
        if (userName) {
            // if name present then the identity is user-assigned managed identity
            // the name option in this case is either client_id or principal_id (object_id) 
            // of the managed identity service principal
            requestOptions.url += `&client_id=${encodeURIComponent(userName)}`;
            if (debug) {
                await logger.logToStderr('Wil try to get token using client_id param...');
            }
        }
        try {
            const accessTokenResponse = await request.get(requestOptions);
            return {
                accessToken: accessTokenResponse.access_token,
                expiresOn: new Date(parseInt(accessTokenResponse.expires_on) * 1000)
            };
        }
        catch (e) {
            if (!userName) {
                throw e;
            }
            // since the userName option can be either client_id or principal_id (object_id) 
            // and the first attempt was using client_id
            // now lets see if the api returned 'not found' response and
            // try to get token using principal_id (object_id)
            let isNotFoundResponse = false;
            if (e.error && e.error.Message) {
                // check if it is Azure Function api 'not found' response
                isNotFoundResponse = (e.error.Message.indexOf("No Managed Identity found") !== -1);
            }
            else if (e.error && e.error.error_description) {
                // check if it is Azure VM api 'not found' response
                isNotFoundResponse = (e.error.error_description === "Identity not found");
            }
            if (!isNotFoundResponse) {
                // it is not a 'not found' response then exit with error
                throw e;
            }
            if (debug) {
                await logger.logToStderr('Wil try to get token using principal_id (also known as object_id) param ...');
            }
            requestOptions.url = requestOptions.url.replace('&client_id=', '&principal_id=');
            requestOptions.headers['x-anonymous'] = true;
            try {
                const accessTokenResponse = await request.get(requestOptions);
                return {
                    accessToken: accessTokenResponse.access_token,
                    expiresOn: new Date(parseInt(accessTokenResponse.expires_on) * 1000)
                };
            }
            catch (err) {
                // will give up and not try any further with the 'msi_res_id' (resource id) query string param
                // since it does not work with the Azure Functions api, but just with the Azure VM api
                if (err.error.code === 'EACCES') {
                    // the CLI does not know if managed identity is actually assigned when EACCES code thrown
                    // so show meaningful message since the raw error response could be misleading 
                    return Promise.reject('Error while logging with Managed Identity. Please check if a Managed Identity is assigned to the current Azure resource.');
                }
                else {
                    throw err;
                }
            }
        }
    }
    async ensureAccessTokenWithFederatedIdentity(resource, logger, debug) {
        if (debug) {
            await logger.logToStderr('Trying to retrieve access token using federated identity...');
        }
        if (process.env.ACTIONS_ID_TOKEN_REQUEST_URL && process.env.ACTIONS_ID_TOKEN_REQUEST_TOKEN) {
            if (debug) {
                await logger.logToStderr('ACTIONS_ID_TOKEN_REQUEST_URL and ACTIONS_ID_TOKEN_REQUEST_TOKEN env variables found. The context is GitHub Actions...');
            }
            const federationToken = await this.getFederationTokenFromGithub(logger, debug);
            return this.getAccessTokenWithFederatedToken(resource, federationToken, logger, debug);
        }
        else if (process.env.SYSTEM_OIDCREQUESTURI) {
            if (debug) {
                await logger.logToStderr('SYSTEM_OIDCREQUESTURI env variable found. The context is Azure DevOps...');
            }
            if (!process.env.SYSTEM_ACCESSTOKEN) {
                throw new CommandError(`The SYSTEM_ACCESSTOKEN environment variable is not available. Please check the Azure DevOps pipeline task configuration. It should contain 'SYSTEM_ACCESSTOKEN: $(System.AccessToken)' in the env section.`);
            }
            const serviceConnectionId = process.env.AZURESUBSCRIPTION_SERVICE_CONNECTION_ID;
            const serviceConnectionAppId = process.env.AZURESUBSCRIPTION_CLIENT_ID;
            const serviceConnectionTenantId = process.env.AZURESUBSCRIPTION_TENANT_ID;
            const useServiceConnection = serviceConnectionId && serviceConnectionAppId && serviceConnectionTenantId;
            if (!useServiceConnection) {
                if (debug) {
                    await logger.logToStderr('Not using a service connection. Run this command in an AzurePowerShell task to be able to use a service connection.');
                }
                if (!this.connection.appId || this.connection.tenant === 'common') {
                    throw new CommandError('The appId and tenant parameters are required when not using a service connection.');
                }
            }
            else {
                if (debug) {
                    if (this.connection.appId || this.connection.tenant !== 'common') {
                        await logger.logToStderr('When using a service connection, the appId and tenant values are updated to the values of the service connection.');
                    }
                    await logger.logToStderr(`Using service connection '${serviceConnectionId}' with app Id '${serviceConnectionAppId}' and tenant Id '${serviceConnectionTenantId}'...`);
                }
                this.connection.appId = serviceConnectionAppId;
                this.connection.tenant = serviceConnectionTenantId;
            }
            const federationToken = await this.getFederationTokenFromAzureDevOps(logger, debug, serviceConnectionId);
            return this.getAccessTokenWithFederatedToken(resource, federationToken, logger, debug);
        }
        else {
            throw new CommandError('Federated identity is currently only supported in GitHub Actions and Azure DevOps.');
        }
    }
    async getFederationTokenFromGithub(logger, debug) {
        if (debug) {
            await logger.logToStderr('Retrieving GitHub federation token...');
        }
        const requestOptions = {
            url: `${process.env.ACTIONS_ID_TOKEN_REQUEST_URL}&audience=${encodeURIComponent('api://AzureADTokenExchange')}`,
            headers: {
                Authorization: `Bearer ${process.env.ACTIONS_ID_TOKEN_REQUEST_TOKEN}`,
                Accept: 'application/json',
                'x-anonymous': true
            },
            responseType: 'json'
        };
        const accessTokenResponse = await request.get(requestOptions);
        return accessTokenResponse.value;
    }
    async getFederationTokenFromAzureDevOps(logger, debug, serviceConnectionId) {
        if (debug) {
            await logger.logToStderr('Retrieving Azure DevOps federation token...');
        }
        const urlSuffix = serviceConnectionId ? `&serviceConnectionId=${serviceConnectionId}` : '';
        const requestOptions = {
            url: `${process.env.SYSTEM_OIDCREQUESTURI}?api-version=7.1${urlSuffix}`,
            headers: {
                Authorization: `Bearer ${process.env.SYSTEM_ACCESSTOKEN}`,
                Accept: 'application/json',
                'Content-Type': 'application/json',
                'x-anonymous': true
            },
            responseType: 'json'
        };
        const accessTokenResponse = await request.post(requestOptions);
        return accessTokenResponse.oidcToken;
    }
    async getAccessTokenWithFederatedToken(resource, federatedToken, logger, debug) {
        if (debug) {
            await logger.logToStderr('Retrieving Entra ID Access Token with federated token...');
        }
        const queryParams = [
            'grant_type=client_credentials',
            `scope=${encodeURIComponent(`${resource}/.default`)}`,
            `client_id=${this.connection.appId}`,
            `client_assertion_type=${encodeURIComponent('urn:ietf:params:oauth:client-assertion-type:jwt-bearer')}`,
            `client_assertion=${federatedToken}`
        ];
        const requestOptions = {
            url: `https://login.microsoftonline.com/${this.connection.tenant}/oauth2/v2.0/token`,
            headers: {
                accept: 'application/json',
                'x-anonymous': true,
                'Content-Type': 'application/x-www-form-urlencoded'
            },
            data: queryParams.join('&'),
            responseType: 'json'
        };
        const accessTokenResponse = await request.post(requestOptions);
        const expiresIn = parseInt(accessTokenResponse.expires_in) * 1000;
        const now = new Date();
        return {
            accessToken: accessTokenResponse.access_token,
            expiresOn: new Date(now.getTime() + expiresIn)
        };
    }
    async ensureAccessTokenWithSecret(resource, logger, debug, fetchNew) {
        this.clientApplication = await this.getConfidentialClient(logger, debug, undefined, undefined, this.connection.secret);
        return this.clientApplication.acquireTokenByClientCredential({
            scopes: [`${resource}/.default`],
            skipCache: fetchNew
        });
    }
    async calculateThumbprint(certificate) {
        const nodeForge = (await import('node-forge')).default;
        const { md, asn1, pki } = nodeForge;
        const messageDigest = md.sha1.create();
        messageDigest.update(asn1.toDer(pki.certificateToAsn1(certificate)).getBytes());
        return messageDigest.digest().toHex();
    }
    static getResourceFromUrl(url) {
        let resource = url;
        const pos = resource.indexOf('/', 8);
        if (pos > -1) {
            resource = resource.substring(0, pos);
        }
        if (resource === 'https://api.bap.microsoft.com' ||
            resource === 'https://api.powerapps.com' ||
            resource.endsWith('.api.bap.microsoft.com')) {
            resource = 'https://service.powerapps.com/';
        }
        if (resource === 'https://api.flow.microsoft.com') {
            resource = 'https://management.azure.com/';
        }
        if (resource === 'https://api.powerbi.com') {
            // api.powerbi.com is not a valid resource
            // we need to use https://analysis.windows.net/powerbi/api instead
            resource = 'https://analysis.windows.net/powerbi/api';
        }
        return resource;
    }
    async getConnectionInfoFromStorage() {
        const tokenStorage = this.getConnectionStorage();
        const json = await tokenStorage.get();
        return JSON.parse(json);
    }
    async storeConnectionInfo() {
        const connectionStorage = this.getConnectionStorage();
        await connectionStorage.set(JSON.stringify(this.connection));
        let allConnections = await this.getAllConnections();
        if (this.connection.active) {
            allConnections = allConnections.filter(c => c.identityId !== this.connection.identityId ||
                c.appId !== this.connection.appId ||
                c.tenant !== this.connection.tenant);
            allConnections.forEach(c => c.active = false);
            allConnections = [{ ...this.connection }, ...allConnections];
        }
        this._allConnections = allConnections;
        const allConnectionsStorage = this.getAllConnectionsStorage();
        await allConnectionsStorage.set(JSON.stringify(allConnections));
    }
    async clearConnectionInfo() {
        const connectionStorage = this.getConnectionStorage();
        const allConnectionsStorage = this.getAllConnectionsStorage();
        await connectionStorage.remove();
        await allConnectionsStorage.remove();
        // we need to manually clear MSAL cache, because MSAL doesn't have support
        // for logging out when using cert-based auth
        const msalCache = this.getMsalCacheStorage();
        await msalCache.remove();
    }
    async removeConnectionInfo(connection, logger, debug) {
        const allConnections = await this.getAllConnections();
        const isCurrentConnection = this.connection.name === connection.name;
        this._allConnections = allConnections.filter(c => c.name !== connection.name);
        // Asserting identityId because it is optional, but required at this point.
        assert(connection.identityId !== undefined);
        // When using an application identity, there is no account in the MSAL TokenCache
        if (this.connection.authType !== AuthType.Certificate &&
            this.connection.authType !== AuthType.Secret &&
            this.connection.authType !== AuthType.Identity &&
            this.connection.authType !== AuthType.FederatedIdentity) {
            this.clientApplication = await this.getPublicClient(logger, debug);
            if (this.clientApplication) {
                const tokenCache = this.clientApplication.getTokenCache();
                const account = await tokenCache.getAccountByLocalId(connection.identityId);
                if (account !== null) {
                    await tokenCache.removeAccount(account);
                }
            }
        }
        const connectionStorage = this.getConnectionStorage();
        const allConnectionsStorage = this.getAllConnectionsStorage();
        await allConnectionsStorage.set(JSON.stringify(this._allConnections));
        if (isCurrentConnection) {
            await connectionStorage.remove();
            this.connection.deactivate();
        }
    }
    getConnectionStorage() {
        return new FileTokenStorage(FileTokenStorage.connectionInfoFilePath());
    }
    getMsalCacheStorage() {
        return new FileTokenStorage(FileTokenStorage.msalCacheFilePath());
    }
    getAllConnectionsStorage() {
        return new FileTokenStorage(FileTokenStorage.allConnectionsFilePath());
    }
    static getEndpointForResource(resource, cloudType) {
        if (Auth.cloudEndpoints[cloudType] &&
            Auth.cloudEndpoints[cloudType][resource]) {
            return Auth.cloudEndpoints[cloudType][resource];
        }
        else {
            return resource;
        }
    }
    async getAllConnectionsFromStorage() {
        const connectionStorage = this.getAllConnectionsStorage();
        const json = await connectionStorage.get();
        return JSON.parse(json);
    }
    async switchToConnection(connection) {
        this.connection.deactivate();
        this._connection = Object.assign(this._connection, connection);
        this._connection.active = true;
        await this.storeConnectionInfo();
    }
    async updateConnection(connection, newName) {
        const allConnections = await this.getAllConnections();
        const existingConnection = allConnections.find(c => c.name === newName);
        const oldName = connection.name;
        if (existingConnection) {
            throw new CommandError(`The connection name '${newName}' is already in use.`);
        }
        connection.name = newName;
        if (this.connection.name === oldName) {
            this._connection.name = newName;
        }
        await this.storeConnectionInfo();
    }
    async getConnection(name) {
        const allConnections = await this.getAllConnections();
        const connection = allConnections.find(i => i.name === name);
        if (!connection) {
            throw new CommandError(`The connection '${name}' cannot be found.`);
        }
        return connection;
    }
    getConnectionDetails(connection) {
        // Asserting name and identityId because they are optional, but required at this point.    
        assert(connection.identityName !== undefined);
        assert(connection.name !== undefined);
        const details = {
            connectionName: connection.name,
            connectedAs: connection.identityName,
            authType: connection.authType,
            appId: connection.appId,
            appTenant: connection.tenant,
            cloudType: CloudType[connection.cloudType]
        };
        return details;
    }
}
Auth.cloudEndpoints = {};
Auth.initialize();
export default new Auth();
//# sourceMappingURL=Auth.js.map