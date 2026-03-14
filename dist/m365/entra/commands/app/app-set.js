var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraAppSetCommand_instances, _a, _EntraAppSetCommand_initTelemetry, _EntraAppSetCommand_initOptions, _EntraAppSetCommand_initValidators, _EntraAppSetCommand_initOptionSets, _EntraAppSetCommand_initTypes;
import fs from 'fs';
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { optionsUtils } from '../../../../utils/optionsUtils.js';
import { entraApp } from '../../../../utils/entraApp.js';
class EntraAppSetCommand extends GraphCommand {
    get name() {
        return commands.APP_SET;
    }
    get description() {
        return 'Updates Entra app registration';
    }
    allowUnknownOptions() {
        return true;
    }
    constructor() {
        super();
        _EntraAppSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraAppSetCommand_instances, "m", _EntraAppSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraAppSetCommand_instances, "m", _EntraAppSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraAppSetCommand_instances, "m", _EntraAppSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraAppSetCommand_instances, "m", _EntraAppSetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _EntraAppSetCommand_instances, "m", _EntraAppSetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            let objectId = await this.getAppObjectId(args, logger);
            objectId = await this.updateUnknownOptions(args, objectId);
            objectId = await this.configureUri(args, objectId, logger);
            objectId = await this.configureRedirectUris(args, objectId, logger);
            objectId = await this.updateAllowPublicClientFlows(args, objectId, logger);
            await this.configureCertificate(args, objectId, logger);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getAppObjectId(args, logger) {
        if (args.options.objectId) {
            return args.options.objectId;
        }
        const { appId, name } = args.options;
        if (this.verbose) {
            await logger.logToStderr(`Retrieving information about Microsoft Entra app ${appId ? appId : name}...`);
        }
        if (appId) {
            const app = await entraApp.getAppRegistrationByAppId(appId, ['id']);
            return app.id;
        }
        else {
            const app = await entraApp.getAppRegistrationByAppName(name, ["id"]);
            return app.id;
        }
    }
    async updateUnknownOptions(args, objectId) {
        const unknownOptions = optionsUtils.getUnknownOptions(args.options, this.options);
        if (Object.keys(unknownOptions).length > 0) {
            const requestBody = {};
            optionsUtils.addUnknownOptionsToPayload(requestBody, unknownOptions);
            const requestOptions = {
                url: `${this.resource}/v1.0/myorganization/applications/${objectId}`,
                headers: {
                    'content-type': 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: requestBody
            };
            await request.patch(requestOptions);
        }
        return objectId;
    }
    async updateAllowPublicClientFlows(args, objectId, logger) {
        if (args.options.allowPublicClientFlows === undefined) {
            return objectId;
        }
        if (this.verbose) {
            await logger.logToStderr(`Configuring Entra application AllowPublicClientFlows option...`);
        }
        const applicationInfo = {
            isFallbackPublicClient: args.options.allowPublicClientFlows
        };
        const requestOptions = {
            url: `${this.resource}/v1.0/myorganization/applications/${objectId}`,
            headers: {
                'content-type': 'application/json;odata.metadata=none'
            },
            responseType: 'json',
            data: applicationInfo
        };
        await request.patch(requestOptions);
        return objectId;
    }
    async configureUri(args, objectId, logger) {
        if (!args.options.uris) {
            return objectId;
        }
        if (this.verbose) {
            await logger.logToStderr(`Configuring Microsoft Entra application ID URI...`);
        }
        const identifierUris = args.options.uris
            .split(',')
            .map(u => u.trim());
        const applicationInfo = {
            identifierUris: identifierUris
        };
        const requestOptions = {
            url: `${this.resource}/v1.0/myorganization/applications/${objectId}`,
            headers: {
                'content-type': 'application/json;odata.metadata=none'
            },
            responseType: 'json',
            data: applicationInfo
        };
        await request.patch(requestOptions);
        return objectId;
    }
    async configureRedirectUris(args, objectId, logger) {
        if (!args.options.redirectUris && !args.options.redirectUrisToRemove) {
            return objectId;
        }
        if (this.verbose) {
            await logger.logToStderr(`Configuring Microsoft Entra application redirect URIs...`);
        }
        const getAppRequestOptions = {
            url: `${this.resource}/v1.0/myorganization/applications/${objectId}`,
            headers: {
                'content-type': 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const application = await request.get(getAppRequestOptions);
        const publicClientRedirectUris = application.publicClient.redirectUris;
        const spaRedirectUris = application.spa.redirectUris;
        const webRedirectUris = application.web.redirectUris;
        // start with existing redirect URIs
        const applicationPatch = {
            publicClient: {
                redirectUris: publicClientRedirectUris
            },
            spa: {
                redirectUris: spaRedirectUris
            },
            web: {
                redirectUris: webRedirectUris
            }
        };
        if (args.options.redirectUrisToRemove) {
            // remove redirect URIs from all platforms
            const redirectUrisToRemove = args.options.redirectUrisToRemove
                .split(',')
                .map(u => u.trim());
            applicationPatch.publicClient.redirectUris =
                publicClientRedirectUris.filter(u => !redirectUrisToRemove.includes(u));
            applicationPatch.spa.redirectUris =
                spaRedirectUris.filter(u => !redirectUrisToRemove.includes(u));
            applicationPatch.web.redirectUris =
                webRedirectUris.filter(u => !redirectUrisToRemove.includes(u));
        }
        if (args.options.redirectUris) {
            const urlsToAdd = args.options.redirectUris
                .split(',')
                .map(u => u.trim());
            // add new redirect URIs. If the URI is already present, it will be ignored
            switch (args.options.platform) {
                case 'spa':
                    applicationPatch.spa.redirectUris
                        .push(...urlsToAdd.filter(u => !spaRedirectUris.includes(u)));
                    break;
                case 'publicClient':
                    applicationPatch.publicClient.redirectUris
                        .push(...urlsToAdd.filter(u => !publicClientRedirectUris.includes(u)));
                    break;
                case 'web':
                    applicationPatch.web.redirectUris
                        .push(...urlsToAdd.filter(u => !webRedirectUris.includes(u)));
            }
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/myorganization/applications/${objectId}`,
            headers: {
                'content-type': 'application/json;odata.metadata=none'
            },
            responseType: 'json',
            data: applicationPatch
        };
        await request.patch(requestOptions);
        return objectId;
    }
    async configureCertificate(args, objectId, logger) {
        if (!args.options.certificateFile && !args.options.certificateBase64Encoded) {
            return;
        }
        if (this.verbose) {
            await logger.logToStderr(`Setting certificate for Microsoft Entra app...`);
        }
        const certificateBase64Encoded = await this.getCertificateBase64Encoded(args, logger);
        const currentKeyCredentials = await this.getCurrentKeyCredentialsList(args, objectId, certificateBase64Encoded, logger);
        if (this.verbose) {
            await logger.logToStderr(`Adding new keyCredential to list`);
        }
        // The KeyCredential graph type defines the 'key' property as 'NullableOption<number>'
        // while it is a base64 encoded string. This is why a cast to any is used here.
        const keyCredentials = currentKeyCredentials.filter(existingCredential => existingCredential.key !== certificateBase64Encoded);
        const newKeyCredential = {
            type: "AsymmetricX509Cert",
            usage: "Verify",
            displayName: args.options.certificateDisplayName,
            key: certificateBase64Encoded
        };
        keyCredentials.push(newKeyCredential);
        return this.updateKeyCredentials(objectId, keyCredentials, logger);
    }
    async getCertificateBase64Encoded(args, logger) {
        if (args.options.certificateBase64Encoded) {
            return args.options.certificateBase64Encoded;
        }
        if (this.debug) {
            await logger.logToStderr(`Reading existing ${args.options.certificateFile}...`);
        }
        try {
            return fs.readFileSync(args.options.certificateFile, { encoding: 'base64' });
        }
        catch (e) {
            throw new Error(`Error reading certificate file: ${e}. Please add the certificate using base64 option '--certificateBase64Encoded'.`);
        }
    }
    // We first retrieve existing certificates because we need to specify the full list of certificates when updating the app.
    async getCurrentKeyCredentialsList(args, objectId, certificateBase64Encoded, logger) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving current keyCredentials list for app`);
        }
        const getAppRequestOptions = {
            url: `${this.resource}/v1.0/myorganization/applications/${objectId}?$select=keyCredentials`,
            headers: {
                'content-type': 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const application = await request.get(getAppRequestOptions);
        return application.keyCredentials || [];
    }
    async updateKeyCredentials(objectId, keyCredentials, logger) {
        if (this.verbose) {
            await logger.logToStderr(`Updating keyCredentials in Microsoft Entra app`);
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/myorganization/applications/${objectId}`,
            headers: {
                'content-type': 'application/json;odata.metadata=none'
            },
            responseType: 'json',
            data: {
                keyCredentials: keyCredentials
            }
        };
        return request.patch(requestOptions);
    }
}
_a = EntraAppSetCommand, _EntraAppSetCommand_instances = new WeakSet(), _EntraAppSetCommand_initTelemetry = function _EntraAppSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            appId: typeof args.options.appId !== 'undefined',
            objectId: typeof args.options.objectId !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            platform: typeof args.options.platform !== 'undefined',
            redirectUris: typeof args.options.redirectUris !== 'undefined',
            redirectUrisToRemove: typeof args.options.redirectUrisToRemove !== 'undefined',
            uris: typeof args.options.uris !== 'undefined',
            certificateFile: typeof args.options.certificateFile !== 'undefined',
            certificateBase64Encoded: typeof args.options.certificateBase64Encoded !== 'undefined',
            certificateDisplayName: typeof args.options.certificateDisplayName !== 'undefined',
            allowPublicClientFlows: typeof args.options.allowPublicClientFlows !== 'undefined'
        });
        this.trackUnknownOptions(this.telemetryProperties, args.options);
    });
}, _EntraAppSetCommand_initOptions = function _EntraAppSetCommand_initOptions() {
    this.options.unshift({ option: '--appId [appId]' }, { option: '--objectId [objectId]' }, { option: '-n, --name [name]' }, { option: '-u, --uris [uris]' }, { option: '-r, --redirectUris [redirectUris]' }, { option: '--certificateFile [certificateFile]' }, { option: '--certificateBase64Encoded [certificateBase64Encoded]' }, { option: '--certificateDisplayName [certificateDisplayName]' }, {
        option: '--platform [platform]',
        autocomplete: _a.aadApplicationPlatform
    }, { option: '--redirectUrisToRemove [redirectUrisToRemove]' }, {
        option: '--allowPublicClientFlows [allowPublicClientFlows]',
        autocomplete: ['true', 'false']
    });
}, _EntraAppSetCommand_initValidators = function _EntraAppSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.certificateFile && args.options.certificateBase64Encoded) {
            return 'Specify either certificateFile or certificateBase64Encoded but not both';
        }
        if (args.options.certificateDisplayName && !args.options.certificateFile && !args.options.certificateBase64Encoded) {
            return 'When you specify certificateDisplayName you also need to specify certificateFile or certificateBase64Encoded';
        }
        if (args.options.certificateFile && !fs.existsSync(args.options.certificateFile)) {
            return 'Certificate file not found';
        }
        if (args.options.redirectUris && !args.options.platform) {
            return `When you specify redirectUris you also need to specify platform`;
        }
        if (args.options.platform &&
            _a.aadApplicationPlatform.indexOf(args.options.platform) < 0) {
            return `${args.options.platform} is not a valid value for platform. Allowed values are ${_a.aadApplicationPlatform.join(', ')}`;
        }
        return true;
    });
}, _EntraAppSetCommand_initOptionSets = function _EntraAppSetCommand_initOptionSets() {
    this.optionSets.push({ options: ['appId', 'objectId', 'name'] });
}, _EntraAppSetCommand_initTypes = function _EntraAppSetCommand_initTypes() {
    this.types.boolean.push('allowPublicClientFlows');
};
EntraAppSetCommand.aadApplicationPlatform = ['spa', 'web', 'publicClient'];
export default new EntraAppSetCommand();
//# sourceMappingURL=app-set.js.map