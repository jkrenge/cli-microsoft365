var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SetupCommand_instances, _SetupCommand_initTelemetry, _SetupCommand_initOptions, _SetupCommand_initValidators;
import chalk from 'chalk';
import os from 'os';
import auth, { AuthType } from '../../Auth.js';
import { cli } from '../../cli/cli.js';
import config from '../../config.js';
import { settingsNames } from '../../settingsNames.js';
import { accessToken } from '../../utils/accessToken.js';
import { entraApp } from '../../utils/entraApp.js';
import { CheckStatus, formatting } from '../../utils/formatting.js';
import { pid } from '../../utils/pid.js';
import { validation } from '../../utils/validation.js';
import AnonymousCommand from '../base/AnonymousCommand.js';
import commands from './commands.js';
import { interactivePreset, powerShellPreset, scriptingPreset } from './setupPresets.js';
export var CliUsageMode;
(function (CliUsageMode) {
    CliUsageMode["Interactively"] = "interactively";
    CliUsageMode["Scripting"] = "scripting";
})(CliUsageMode || (CliUsageMode = {}));
export var CliExperience;
(function (CliExperience) {
    CliExperience["Beginner"] = "beginner";
    CliExperience["Proficient"] = "proficient";
})(CliExperience || (CliExperience = {}));
export var EntraAppConfig;
(function (EntraAppConfig) {
    EntraAppConfig["Create"] = "create";
    EntraAppConfig["UseExisting"] = "useExisting";
    EntraAppConfig["Skip"] = "skip";
})(EntraAppConfig || (EntraAppConfig = {}));
export var NewEntraAppScopes;
(function (NewEntraAppScopes) {
    NewEntraAppScopes["Minimal"] = "minimal";
    NewEntraAppScopes["All"] = "all";
})(NewEntraAppScopes || (NewEntraAppScopes = {}));
export var HelpMode;
(function (HelpMode) {
    HelpMode["Full"] = "full";
    HelpMode["Options"] = "options";
})(HelpMode || (HelpMode = {}));
class SetupCommand extends AnonymousCommand {
    get name() {
        return commands.SETUP;
    }
    get description() {
        return 'Sets up CLI for Microsoft 365 based on your preferences';
    }
    constructor() {
        super();
        _SetupCommand_instances.add(this);
        __classPrivateFieldGet(this, _SetupCommand_instances, "m", _SetupCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SetupCommand_instances, "m", _SetupCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SetupCommand_instances, "m", _SetupCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        let settings;
        if (args.options.interactive || args.options.scripting) {
            settings = {};
            if (args.options.interactive) {
                Object.assign(settings, interactivePreset);
            }
            else if (args.options.scripting) {
                Object.assign(settings, scriptingPreset);
            }
            if (pid.isPowerShell()) {
                Object.assign(settings, powerShellPreset);
            }
            await this.configureSettings({ preferences: {}, settings, silent: true, logger });
            return;
        }
        await logger.logToStderr(`Welcome to the CLI for Microsoft 365 setup!`);
        await logger.logToStderr(`This command will guide you through the process of configuring the CLI for your needs.`);
        await logger.logToStderr(`Please, answer the following questions and we'll define a set of settings to best match how you intend to use the CLI.`);
        await logger.logToStderr('');
        const preferences = {};
        if (!args.options.skipApp) {
            const entraAppConfig = {
                message: 'CLI for Microsoft 365 requires a Microsoft Entra app. Do you want to create a new app registration or use an existing one?',
                choices: [
                    { name: 'Create a new app registration', value: EntraAppConfig.Create },
                    { name: 'Use an existing app registration', value: EntraAppConfig.UseExisting },
                    { name: 'Skip configuring app registration', value: EntraAppConfig.Skip }
                ]
            };
            preferences.entraApp = await cli.promptForSelection(entraAppConfig);
            switch (preferences.entraApp) {
                case EntraAppConfig.Create: {
                    const newEntraAppScopesConfig = {
                        message: 'What scopes should the new app registration have?',
                        choices: [
                            { name: 'User.Read (you will need to add the necessary permissions yourself)', value: NewEntraAppScopes.Minimal },
                            { name: 'All (easy way to use all CLI commands)', value: NewEntraAppScopes.All }
                        ]
                    };
                    preferences.newEntraAppScopes = await cli.promptForSelection(newEntraAppScopesConfig);
                    break;
                }
                case EntraAppConfig.UseExisting: {
                    const existingApp = await this.configureExistingEntraApp(logger);
                    Object.assign(preferences, existingApp);
                    break;
                }
            }
        }
        else {
            preferences.entraApp = EntraAppConfig.Skip;
        }
        const usageModeConfig = {
            message: 'How do you plan to use the CLI?',
            choices: [
                { name: 'Interactively', value: CliUsageMode.Interactively },
                { name: 'Scripting', value: CliUsageMode.Scripting }
            ]
        };
        preferences.usageMode = await cli.promptForSelection(usageModeConfig);
        if (preferences.usageMode === CliUsageMode.Scripting) {
            const usedInPowerShellConfig = {
                message: 'Are you going to use the CLI in PowerShell?',
                default: pid.isPowerShell()
            };
            preferences.usedInPowerShell = await cli.promptForConfirmation(usedInPowerShellConfig);
        }
        const experienceConfig = {
            message: 'How experienced are you in using the CLI?',
            choices: [
                { name: 'Beginner', value: CliExperience.Beginner },
                { name: 'Proficient', value: CliExperience.Proficient }
            ]
        };
        preferences.experience = await cli.promptForSelection(experienceConfig);
        const summaryConfig = {
            message: this.getSummaryMessage(preferences)
        };
        preferences.summary = await cli.promptForConfirmation(summaryConfig);
        if (!preferences.summary) {
            return;
        }
        // used only for testing. Normally, we'd get the settings from the answers
        /* c8 ignore next 3 */
        if (!settings) {
            settings = this.getSettings(preferences);
        }
        await logger.logToStderr('');
        await logger.logToStderr('Configuring settings...');
        await logger.logToStderr('');
        await this.configureSettings({ preferences, settings, silent: false, logger });
        if (!this.verbose) {
            await logger.logToStderr('');
            await logger.logToStderr(chalk.green('DONE'));
        }
    }
    async configureExistingEntraApp(logger) {
        await logger.logToStderr('Please provide the details of the existing app registration.');
        let clientCertificateFile;
        let clientCertificateBase64Encoded;
        let clientCertificatePassword;
        const clientId = await cli.promptForInput({
            message: 'Client ID:',
            /* c8 ignore next */
            validate: value => validation.isValidGuid(value) ? true : 'The specified value is not a valid GUID.'
        });
        const tenantId = await cli.promptForInput({
            message: 'Tenant ID (leave common if the app is multitenant):',
            default: 'common',
            /* c8 ignore next */
            validate: value => value === 'common' || validation.isValidGuid(value) ? true : `Tenant ID must be a valid GUID or 'common'.`
        });
        const clientSecret = await cli.promptForInput({
            message: 'Client secret (leave empty if you use a certificate or a public client):'
        });
        if (!clientSecret) {
            clientCertificateFile = await cli.promptForInput({
                message: `Path to the client certificate file (leave empty if you want to specify a base64-encoded certificate string):`
            });
            if (!clientCertificateFile) {
                clientCertificateBase64Encoded = await cli.promptForInput({
                    message: `Base64-encoded certificate string (leave empty if you don't connect using a certificate):`
                });
            }
            if (clientCertificateFile || clientCertificateBase64Encoded) {
                clientCertificatePassword = await cli.promptForInput({
                    message: 'Password for the client certificate (leave empty if the certificate is not password-protected):'
                });
            }
        }
        return {
            clientId,
            tenantId,
            clientSecret,
            clientCertificateFile,
            clientCertificateBase64Encoded,
            clientCertificatePassword
        };
    }
    async createNewEntraApp(preferences, logger) {
        if (!await cli.promptForConfirmation({
            message: 'CLI for Microsoft 365 will now sign in to your Microsoft 365 tenant as Microsoft Azure CLI to create a new app registration. Continue?',
            default: false
        })) {
            throw 'Cancelled';
        }
        // setup auth
        auth.connection.authType = AuthType.Browser;
        // Microsoft Azure CLI app ID
        auth.connection.appId = '04b07795-8ddb-461a-bbee-02f9e1bf7b46';
        auth.connection.tenant = 'common';
        await auth.ensureAccessToken(auth.defaultResource, logger, this.debug);
        auth.connection.active = true;
        const options = {
            allowPublicClientFlows: true,
            apisDelegated: (preferences.newEntraAppScopes === NewEntraAppScopes.All ? config.allScopes : config.minimalScopes).join(','),
            implicitFlow: false,
            multitenant: false,
            name: 'CLI for M365',
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
        return appInfo;
    }
    getSummaryMessage(preferences) {
        const messageLines = [`Based on your preferences, we'll configure the following settings:`];
        switch (preferences.entraApp) {
            case EntraAppConfig.Create:
                messageLines.push(`- Entra app: Create a new app registration with ${preferences.newEntraAppScopes} scopes`);
                break;
            case EntraAppConfig.UseExisting:
                messageLines.push(`- Entra app: use existing`);
                messageLines.push(`  - Client ID: ${preferences.clientId}`);
                messageLines.push(`  - Tenant ID: ${preferences.tenantId}`);
                if (preferences.clientSecret) {
                    messageLines.push(`  - Client secret: ${preferences.clientSecret}`);
                }
                if (preferences.clientCertificateFile) {
                    messageLines.push(`  - Client certificate file: ${preferences.clientCertificateFile}`);
                }
                if (preferences.clientCertificateBase64Encoded) {
                    messageLines.push(`  - Client certificate base64-encoded: ${preferences.clientCertificateBase64Encoded}`);
                }
                if (preferences.clientCertificatePassword) {
                    messageLines.push(`  - Client certificate password: ${preferences.clientCertificatePassword}`);
                }
                break;
            case EntraAppConfig.Skip:
                messageLines.push(`- Entra app: skip`);
                break;
        }
        const settings = this.getSettings(preferences);
        for (const [key, value] of Object.entries(settings)) {
            messageLines.push(`- ${key}: ${value}`);
        }
        messageLines.push('', 'You can change any of these settings later using the `m365 cli config set` command or reset them to default using `m365 cli config reset`.', '', 'Do you want to apply these settings now?');
        return messageLines.join(os.EOL);
    }
    getSettings(answers) {
        const settings = {};
        switch (answers.usageMode) {
            case CliUsageMode.Interactively:
                Object.assign(settings, interactivePreset);
                break;
            case CliUsageMode.Scripting:
                Object.assign(settings, scriptingPreset);
                break;
        }
        if (answers.usedInPowerShell === true) {
            Object.assign(settings, powerShellPreset);
        }
        switch (answers.experience) {
            case CliExperience.Beginner:
                settings.helpMode = HelpMode.Full;
                break;
            case CliExperience.Proficient:
                settings.helpMode = HelpMode.Options;
                break;
        }
        switch (answers.entraApp) {
            case EntraAppConfig.Create:
                settings.authType = 'browser';
                break;
            case EntraAppConfig.UseExisting:
                if (answers.clientSecret) {
                    settings.authType = 'secret';
                    break;
                }
                if (answers.clientCertificateFile || answers.clientCertificateBase64Encoded) {
                    settings.authType = 'certificate';
                    break;
                }
                settings.authType = 'browser';
                break;
        }
        return settings;
    }
    async configureSettings({ preferences, settings, silent, logger }) {
        switch (preferences.entraApp) {
            case EntraAppConfig.Create: {
                if (this.verbose) {
                    await logger.logToStderr('Creating a new Entra app...');
                }
                const appSettings = await this.createNewEntraApp(preferences, logger);
                Object.assign(settings, {
                    clientId: appSettings.appId,
                    tenantId: appSettings.tenantId
                });
                cli.getConfig().delete(settingsNames.clientSecret);
                cli.getConfig().delete(settingsNames.clientCertificateFile);
                cli.getConfig().delete(settingsNames.clientCertificateBase64Encoded);
                cli.getConfig().delete(settingsNames.clientCertificatePassword);
                break;
            }
            case EntraAppConfig.UseExisting:
                Object.assign(settings, {
                    clientId: preferences.clientId,
                    tenantId: preferences.tenantId,
                    clientSecret: preferences.clientSecret,
                    clientCertificateFile: preferences.clientCertificateFile,
                    clientCertificateBase64Encoded: preferences.clientCertificateBase64Encoded,
                    clientCertificatePassword: preferences.clientCertificatePassword
                });
                break;
            case EntraAppConfig.Skip:
                break;
        }
        if (this.debug) {
            await logger.logToStderr('Configuring settings...');
            await logger.logToStderr(JSON.stringify(settings, null, 2));
        }
        for (const [key, value] of Object.entries(settings)) {
            cli.getConfig().set(key, value);
            if (!silent) {
                await logger.logToStderr(formatting.getStatus(CheckStatus.Success, `${key}: ${value}`));
            }
        }
    }
}
_SetupCommand_instances = new WeakSet(), _SetupCommand_initTelemetry = function _SetupCommand_initTelemetry() {
    this.telemetry.push((args) => {
        const properties = {
            interactive: !!args.options.interactive,
            scripting: !!args.options.scripting,
            skipApp: !!args.options.skipApp
        };
        Object.assign(this.telemetryProperties, properties);
    });
}, _SetupCommand_initOptions = function _SetupCommand_initOptions() {
    this.options.unshift({ option: '--interactive' }, { option: '--scripting' }, { option: '--skipApp' });
}, _SetupCommand_initValidators = function _SetupCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.interactive && args.options.scripting) {
            return 'Specify either interactive or scripting but not both';
        }
        return true;
    });
};
export default new SetupCommand();
//# sourceMappingURL=setup.js.map