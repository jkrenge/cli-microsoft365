var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraAppAddCommand_instances, _a, _EntraAppAddCommand_initTelemetry, _EntraAppAddCommand_initOptions, _EntraAppAddCommand_initValidators, _EntraAppAddCommand_initOptionSets;
import fs from 'fs';
import { v4 } from 'uuid';
import auth from '../../../../Auth.js';
import request from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { entraApp } from '../../../../utils/entraApp.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { optionsUtils } from '../../../../utils/optionsUtils.js';
class EntraAppAddCommand extends GraphCommand {
    get name() {
        return commands.APP_ADD;
    }
    get description() {
        return 'Creates new Entra app registration';
    }
    allowUnknownOptions() {
        return true;
    }
    constructor() {
        super();
        _EntraAppAddCommand_instances.add(this);
        this.appName = '';
        __classPrivateFieldGet(this, _EntraAppAddCommand_instances, "m", _EntraAppAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraAppAddCommand_instances, "m", _EntraAppAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraAppAddCommand_instances, "m", _EntraAppAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraAppAddCommand_instances, "m", _EntraAppAddCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (!args.options.name && this.manifest) {
            args.options.name = this.manifest.name;
        }
        this.appName = args.options.name;
        try {
            const apis = await entraApp.resolveApis({
                options: args.options,
                manifest: this.manifest,
                logger,
                verbose: this.verbose,
                debug: this.debug
            });
            let appInfo = await entraApp.createAppRegistration({
                options: args.options,
                unknownOptions: optionsUtils.getUnknownOptions(args.options, this.options),
                apis,
                logger,
                verbose: this.verbose,
                debug: this.debug
            });
            // based on the assumption that we're adding Microsoft Entra app to the current
            // directory. If we in the future extend the command with allowing
            // users to create Microsoft Entra app in a different directory, we'll need to
            // adjust this
            appInfo.tenantId = accessToken.getTenantIdFromAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken);
            appInfo = await this.updateAppFromManifest(args, appInfo);
            appInfo = await entraApp.grantAdminConsent({
                appInfo,
                appPermissions: entraApp.appPermissions,
                adminConsent: args.options.grantAdminConsent,
                logger,
                debug: this.debug
            });
            appInfo = await this.configureUri(args, appInfo, logger);
            appInfo = await this.configureSecret(args, appInfo, logger);
            const _appInfo = await this.saveAppInfo(args, appInfo, logger);
            appInfo = {
                appId: _appInfo.appId,
                objectId: _appInfo.id,
                tenantId: _appInfo.tenantId
            };
            if (_appInfo.secrets) {
                appInfo.secrets = _appInfo.secrets;
            }
            await logger.log(appInfo);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async configureSecret(args, appInfo, logger) {
        if (!args.options.withSecret || (appInfo.secrets && appInfo.secrets.length > 0)) {
            return appInfo;
        }
        if (this.verbose) {
            await logger.logToStderr(`Configure Microsoft Entra app secret...`);
        }
        const secret = await this.createSecret({ appObjectId: appInfo.id });
        if (!appInfo.secrets) {
            appInfo.secrets = [];
        }
        appInfo.secrets.push(secret);
        return appInfo;
    }
    async createSecret({ appObjectId, displayName = undefined, expirationDate = undefined }) {
        let secretExpirationDate = expirationDate;
        if (!secretExpirationDate) {
            secretExpirationDate = new Date();
            secretExpirationDate.setFullYear(secretExpirationDate.getFullYear() + 1);
        }
        const secretName = displayName ?? 'Default';
        const requestOptions = {
            url: `${this.resource}/v1.0/myorganization/applications/${appObjectId}/addPassword`,
            headers: {
                'content-type': 'application/json'
            },
            responseType: 'json',
            data: {
                passwordCredential: {
                    displayName: secretName,
                    endDateTime: secretExpirationDate.toISOString()
                }
            }
        };
        const response = await request.post(requestOptions);
        return {
            displayName: secretName,
            value: response.secretText
        };
    }
    async updateAppFromManifest(args, appInfo) {
        if (!args.options.manifest) {
            return appInfo;
        }
        const v2Manifest = JSON.parse(args.options.manifest);
        // remove properties that might be coming from the original app that was
        // used to create the manifest and which can't be updated
        delete v2Manifest.id;
        delete v2Manifest.appId;
        delete v2Manifest.publisherDomain;
        // extract secrets from the manifest. Store them in a separate variable
        const secrets = this.getSecretsFromManifest(v2Manifest);
        // Azure Portal returns v2 manifest whereas the Graph API expects a v1.6
        if (args.options.apisApplication || args.options.apisDelegated) {
            // take submitted delegated / application permissions as options
            // otherwise, they will be skipped in the app update
            v2Manifest.requiredResourceAccess = appInfo.requiredResourceAccess;
        }
        if (args.options.redirectUris) {
            // take submitted redirectUris/platform as options
            // otherwise, they will be removed from the app
            v2Manifest.replyUrlsWithType = args.options.redirectUris.split(',').map(u => {
                return {
                    url: u.trim(),
                    type: this.translatePlatformToType(args.options.platform)
                };
            });
        }
        if (args.options.multitenant) {
            // override manifest setting when using multitenant flag
            v2Manifest.signInAudience = 'AzureADMultipleOrgs';
        }
        if (args.options.implicitFlow) {
            // remove manifest settings when using implicitFlow flag
            delete v2Manifest.oauth2AllowIdTokenImplicitFlow;
            delete v2Manifest.oauth2AllowImplicitFlow;
        }
        if (args.options.scopeName) {
            // override manifest setting when using options.
            delete v2Manifest.oauth2Permissions;
        }
        if (args.options.certificateFile || args.options.certificateBase64Encoded) {
            // override manifest setting when using options.
            delete v2Manifest.keyCredentials;
        }
        const graphManifest = this.transformManifest(v2Manifest);
        const updateAppRequestOptions = {
            url: `${this.resource}/v1.0/myorganization/applications/${appInfo.id}`,
            headers: {
                'content-type': 'application/json'
            },
            responseType: 'json',
            data: graphManifest
        };
        await request.patch(updateAppRequestOptions);
        await this.updatePreAuthorizedAppsFromManifest(v2Manifest, appInfo);
        await this.createSecrets(secrets, appInfo);
        return appInfo;
    }
    getSecretsFromManifest(manifest) {
        if (!manifest.passwordCredentials || manifest.passwordCredentials.length === 0) {
            return [];
        }
        const secrets = manifest.passwordCredentials.map((c) => {
            const startDate = new Date(c.startDate);
            const endDate = new Date(c.endDate);
            const expirationDate = new Date();
            expirationDate.setMilliseconds(endDate.valueOf() - startDate.valueOf());
            return {
                name: c.displayName,
                expirationDate
            };
        });
        // delete the secrets from the manifest so that we won't try to set them
        // from the manifest
        delete manifest.passwordCredentials;
        return secrets;
    }
    async updatePreAuthorizedAppsFromManifest(manifest, appInfo) {
        if (!manifest ||
            !manifest.preAuthorizedApplications ||
            manifest.preAuthorizedApplications.length === 0) {
            return appInfo;
        }
        const graphManifest = {
            api: {
                preAuthorizedApplications: manifest.preAuthorizedApplications
            }
        };
        graphManifest.api.preAuthorizedApplications.forEach((p) => {
            p.delegatedPermissionIds = p.permissionIds;
            delete p.permissionIds;
        });
        const updateAppRequestOptions = {
            url: `${this.resource}/v1.0/myorganization/applications/${appInfo.id}`,
            headers: {
                'content-type': 'application/json'
            },
            responseType: 'json',
            data: graphManifest
        };
        await request.patch(updateAppRequestOptions);
        return appInfo;
    }
    async createSecrets(secrets, appInfo) {
        if (secrets.length === 0) {
            return appInfo;
        }
        const secretsOutput = await Promise
            .all(secrets.map(secret => this.createSecret({
            appObjectId: appInfo.id,
            displayName: secret.name,
            expirationDate: secret.expirationDate
        })));
        appInfo.secrets = secretsOutput;
        return appInfo;
    }
    transformManifest(v2Manifest) {
        const graphManifest = JSON.parse(JSON.stringify(v2Manifest));
        // add missing properties
        if (!graphManifest.api) {
            graphManifest.api = {};
        }
        if (!graphManifest.info) {
            graphManifest.info = {};
        }
        if (!graphManifest.web) {
            graphManifest.web = {
                implicitGrantSettings: {},
                redirectUris: []
            };
        }
        if (!graphManifest.spa) {
            graphManifest.spa = {
                redirectUris: []
            };
        }
        // remove properties that have no equivalent in v1.6
        const unsupportedProperties = [
            'accessTokenAcceptedVersion',
            'disabledByMicrosoftStatus',
            'errorUrl',
            'oauth2RequirePostResponse',
            'oauth2AllowUrlPathMatching',
            'orgRestrictions',
            'samlMetadataUrl'
        ];
        unsupportedProperties.forEach(p => delete graphManifest[p]);
        graphManifest.api.acceptMappedClaims = v2Manifest.acceptMappedClaims;
        delete graphManifest.acceptMappedClaims;
        graphManifest.isFallbackPublicClient = v2Manifest.allowPublicClient;
        delete graphManifest.allowPublicClient;
        graphManifest.info.termsOfServiceUrl = v2Manifest.informationalUrls?.termsOfService;
        graphManifest.info.supportUrl = v2Manifest.informationalUrls?.support;
        graphManifest.info.privacyStatementUrl = v2Manifest.informationalUrls?.privacy;
        graphManifest.info.marketingUrl = v2Manifest.informationalUrls?.marketing;
        delete graphManifest.informationalUrls;
        graphManifest.api.knownClientApplications = v2Manifest.knownClientApplications;
        delete graphManifest.knownClientApplications;
        graphManifest.info.logoUrl = v2Manifest.logoUrl;
        delete graphManifest.logoUrl;
        graphManifest.web.logoutUrl = v2Manifest.logoutUrl;
        delete graphManifest.logoutUrl;
        graphManifest.displayName = v2Manifest.name;
        delete graphManifest.name;
        graphManifest.web.implicitGrantSettings.enableAccessTokenIssuance = v2Manifest.oauth2AllowImplicitFlow;
        delete graphManifest.oauth2AllowImplicitFlow;
        graphManifest.web.implicitGrantSettings.enableIdTokenIssuance = v2Manifest.oauth2AllowIdTokenImplicitFlow;
        delete graphManifest.oauth2AllowIdTokenImplicitFlow;
        graphManifest.api.oauth2PermissionScopes = v2Manifest.oauth2Permissions;
        delete graphManifest.oauth2Permissions;
        if (graphManifest.api.oauth2PermissionScopes) {
            graphManifest.api.oauth2PermissionScopes.forEach((scope) => {
                delete scope.lang;
                delete scope.origin;
            });
        }
        delete graphManifest.oauth2RequiredPostResponse;
        // MS Graph doesn't support creating OAuth2 permissions and pre-authorized
        // apps in one request. This is why we need to remove it here and do it in
        // the next request
        delete graphManifest.preAuthorizedApplications;
        if (v2Manifest.replyUrlsWithType) {
            v2Manifest.replyUrlsWithType.forEach((urlWithType) => {
                if (urlWithType.type === 'Web') {
                    graphManifest.web.redirectUris.push(urlWithType.url);
                    return;
                }
                if (urlWithType.type === 'Spa') {
                    graphManifest.spa.redirectUris.push(urlWithType.url);
                    return;
                }
            });
            delete graphManifest.replyUrlsWithType;
        }
        graphManifest.web.homePageUrl = v2Manifest.signInUrl;
        delete graphManifest.signInUrl;
        if (graphManifest.appRoles) {
            graphManifest.appRoles.forEach((role) => {
                delete role.lang;
            });
        }
        return graphManifest;
    }
    async configureUri(args, appInfo, logger) {
        if (!args.options.uri) {
            return appInfo;
        }
        if (this.verbose) {
            await logger.logToStderr(`Configuring Microsoft Entra application ID URI...`);
        }
        const applicationInfo = {};
        if (args.options.uri) {
            const appUri = args.options.uri.replace(/_appId_/g, appInfo.appId);
            applicationInfo.identifierUris = [appUri];
        }
        if (args.options.scopeName) {
            applicationInfo.api = {
                oauth2PermissionScopes: [{
                        adminConsentDescription: args.options.scopeAdminConsentDescription,
                        adminConsentDisplayName: args.options.scopeAdminConsentDisplayName,
                        id: v4(),
                        type: args.options.scopeConsentBy === 'adminsAndUsers' ? 'User' : 'Admin',
                        value: args.options.scopeName
                    }]
            };
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/myorganization/applications/${appInfo.id}`,
            headers: {
                'content-type': 'application/json;odata.metadata=none'
            },
            responseType: 'json',
            data: applicationInfo
        };
        await request.patch(requestOptions);
        return appInfo;
    }
    async saveAppInfo(args, appInfo, logger) {
        if (!args.options.save) {
            return appInfo;
        }
        const filePath = '.m365rc.json';
        if (this.verbose) {
            await logger.logToStderr(`Saving Microsoft Entra app registration information to the ${filePath} file...`);
        }
        let m365rc = {};
        if (fs.existsSync(filePath)) {
            if (this.debug) {
                await logger.logToStderr(`Reading existing ${filePath}...`);
            }
            try {
                const fileContents = fs.readFileSync(filePath, 'utf8');
                if (fileContents) {
                    m365rc = JSON.parse(fileContents);
                }
            }
            catch (e) {
                await logger.logToStderr(`Error reading ${filePath}: ${e}. Please add app info to ${filePath} manually.`);
                return Promise.resolve(appInfo);
            }
        }
        if (!m365rc.apps) {
            m365rc.apps = [];
        }
        m365rc.apps.push({
            appId: appInfo.appId,
            name: this.appName
        });
        try {
            fs.writeFileSync(filePath, JSON.stringify(m365rc, null, 2));
        }
        catch (e) {
            await logger.logToStderr(`Error writing ${filePath}: ${e}. Please add app info to ${filePath} manually.`);
        }
        return Promise.resolve(appInfo);
    }
    translatePlatformToType(platform) {
        if (platform === 'publicClient') {
            return 'InstalledClient';
        }
        return platform.charAt(0).toUpperCase() + platform.substring(1);
    }
}
_a = EntraAppAddCommand, _EntraAppAddCommand_instances = new WeakSet(), _EntraAppAddCommand_initTelemetry = function _EntraAppAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            apis: typeof args.options.apisDelegated !== 'undefined',
            implicitFlow: args.options.implicitFlow,
            multitenant: args.options.multitenant,
            platform: args.options.platform,
            redirectUris: typeof args.options.redirectUris !== 'undefined',
            scopeAdminConsentDescription: typeof args.options.scopeAdminConsentDescription !== 'undefined',
            scopeAdminConsentDisplayName: typeof args.options.scopeAdminConsentDisplayName !== 'undefined',
            scopeConsentBy: args.options.scopeConsentBy,
            scopeName: typeof args.options.scopeName !== 'undefined',
            uri: typeof args.options.uri !== 'undefined',
            withSecret: args.options.withSecret,
            certificateFile: typeof args.options.certificateFile !== 'undefined',
            certificateBase64Encoded: typeof args.options.certificateBase64Encoded !== 'undefined',
            certificateDisplayName: typeof args.options.certificateDisplayName !== 'undefined',
            grantAdminConsent: typeof args.options.grantAdminConsent !== 'undefined',
            allowPublicClientFlows: typeof args.options.allowPublicClientFlows !== 'undefined',
            bundleId: typeof args.options.bundleId !== 'undefined',
            signatureHash: typeof args.options.signatureHash !== 'undefined'
        });
    });
}, _EntraAppAddCommand_initOptions = function _EntraAppAddCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name [name]'
    }, {
        option: '--multitenant'
    }, {
        option: '-r, --redirectUris [redirectUris]'
    }, {
        option: '-p, --platform [platform]',
        autocomplete: _a.entraApplicationPlatform
    }, {
        option: '--implicitFlow'
    }, {
        option: '-s, --withSecret'
    }, {
        option: '--apisDelegated [apisDelegated]'
    }, {
        option: '--apisApplication [apisApplication]'
    }, {
        option: '-u, --uri [uri]'
    }, {
        option: '--scopeName [scopeName]'
    }, {
        option: '--scopeConsentBy [scopeConsentBy]',
        autocomplete: _a.entraAppScopeConsentBy
    }, {
        option: '--scopeAdminConsentDisplayName [scopeAdminConsentDisplayName]'
    }, {
        option: '--scopeAdminConsentDescription [scopeAdminConsentDescription]'
    }, {
        option: '--certificateFile [certificateFile]'
    }, {
        option: '--certificateBase64Encoded [certificateBase64Encoded]'
    }, {
        option: '--certificateDisplayName [certificateDisplayName]'
    }, {
        option: '--manifest [manifest]'
    }, {
        option: '--bundleId [bundleId]'
    }, {
        option: '--signatureHash [signatureHash]'
    }, {
        option: '--save'
    }, {
        option: '--grantAdminConsent'
    }, {
        option: '--allowPublicClientFlows'
    });
}, _EntraAppAddCommand_initValidators = function _EntraAppAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.platform &&
            _a.entraApplicationPlatform.indexOf(args.options.platform) < 0) {
            return `${args.options.platform} is not a valid value for platform. Allowed values are ${_a.entraApplicationPlatform.join(', ')}`;
        }
        if (args.options.redirectUris && !args.options.platform) {
            return `When you specify redirectUris you also need to specify platform`;
        }
        if (args.options.platform && ['spa', 'web', 'publicClient'].indexOf(args.options.platform) > -1 && !args.options.redirectUris) {
            return `When you use platform spa, web or publicClient, you'll need to specify redirectUris`;
        }
        if (args.options.certificateFile && args.options.certificateBase64Encoded) {
            return 'Specify either certificateFile or certificateBase64Encoded but not both';
        }
        if (args.options.certificateDisplayName && !args.options.certificateFile && !args.options.certificateBase64Encoded) {
            return 'When you specify certificateDisplayName you also need to specify certificateFile or certificateBase64Encoded';
        }
        if (args.options.certificateFile && !fs.existsSync(args.options.certificateFile)) {
            return 'Certificate file not found';
        }
        if (args.options.scopeName) {
            if (!args.options.uri) {
                return `When you specify scopeName you also need to specify uri`;
            }
            if (!args.options.scopeAdminConsentDescription) {
                return `When you specify scopeName you also need to specify scopeAdminConsentDescription`;
            }
            if (!args.options.scopeAdminConsentDisplayName) {
                return `When you specify scopeName you also need to specify scopeAdminConsentDisplayName`;
            }
        }
        if (args.options.scopeConsentBy &&
            _a.entraAppScopeConsentBy.indexOf(args.options.scopeConsentBy) < 0) {
            return `${args.options.scopeConsentBy} is not a valid value for scopeConsentBy. Allowed values are ${_a.entraAppScopeConsentBy.join(', ')}`;
        }
        if (args.options.manifest) {
            try {
                this.manifest = JSON.parse(args.options.manifest);
                if (!args.options.name && !this.manifest.name) {
                    return `Specify the name of the app to create either through the 'name' option or the 'name' property in the manifest`;
                }
            }
            catch (e) {
                return `Error while parsing the specified manifest: ${e}`;
            }
        }
        if (args.options.platform === 'apple' && !args.options.bundleId) {
            return `When you use platform apple, you'll need to specify bundleId`;
        }
        if (args.options.platform === 'android' && (!args.options.bundleId || !args.options.signatureHash)) {
            return `When you use platform android, you'll need to specify bundleId and signatureHash`;
        }
        return true;
    });
}, _EntraAppAddCommand_initOptionSets = function _EntraAppAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['name', 'manifest'] });
};
EntraAppAddCommand.entraApplicationPlatform = ['spa', 'web', 'publicClient', 'apple', 'android'];
EntraAppAddCommand.entraAppScopeConsentBy = ['admins', 'adminsAndUsers'];
export default new EntraAppAddCommand();
//# sourceMappingURL=app-add.js.map