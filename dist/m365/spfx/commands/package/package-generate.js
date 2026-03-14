var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpfxPackageGenerateCommand_instances, _a, _SpfxPackageGenerateCommand_initTelemetry, _SpfxPackageGenerateCommand_initOptions, _SpfxPackageGenerateCommand_initValidators;
import AdmZip from 'adm-zip';
import fs from 'fs';
import os from 'os';
import path from 'path';
import url from 'url';
import { v4 } from 'uuid';
import { fsUtil } from '../../../../utils/fsUtil.js';
import AnonymousCommand from '../../../base/AnonymousCommand.js';
import commands from '../../commands.js';
const __dirname = url.fileURLToPath(new URL('.', import.meta.url));
class SpfxPackageGenerateCommand extends AnonymousCommand {
    get name() {
        return commands.PACKAGE_GENERATE;
    }
    get description() {
        return 'Generates SharePoint Framework solution package with a no-framework web part rendering the specified HTML snippet';
    }
    constructor() {
        super();
        _SpfxPackageGenerateCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpfxPackageGenerateCommand_instances, "m", _SpfxPackageGenerateCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpfxPackageGenerateCommand_instances, "m", _SpfxPackageGenerateCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpfxPackageGenerateCommand_instances, "m", _SpfxPackageGenerateCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const supportedHosts = ['SharePointWebPart'];
        if (args.options.enableForTeams === 'tab' || args.options.enableForTeams === 'all') {
            supportedHosts.push('TeamsTab');
        }
        if (args.options.enableForTeams === 'personalApp' || args.options.enableForTeams === 'all') {
            supportedHosts.push('TeamsPersonalApp');
        }
        const tokens = {
            clientSideAssetsFeatureId: _a.generateNewId(),
            developerName: args.options.developerName || 'Contoso',
            developerWebsiteUrl: args.options.developerWebsiteUrl || 'https://contoso.com/my-app',
            developerPrivacyUrl: args.options.developerPrivacyUrl || 'https://contoso.com/privacy',
            developerTermsOfUseUrl: args.options.developerTermsOfUseUrl || 'https://contoso.com/terms-of-use',
            developerMpnId: args.options.developerMpnId || '000000',
            exposePageContextGlobally: args.options.exposePageContextGlobally ? '!0' : '!1',
            exposeTeamsContextGlobally: args.options.exposeTeamsContextGlobally ? '!0' : '!1',
            html: args.options.html.replace(/"/g, '\\"').replace(/\r\n/g, ' ').replace(/\n/g, ' '),
            packageName: _a.getSafePackageName(args.options.webPartTitle),
            productId: _a.generateNewId(),
            skipFeatureDeployment: (args.options.allowTenantWideDeployment === true).toString(),
            supportedHosts: JSON.stringify(supportedHosts).replace(/"/g, '&quot;'),
            webPartId: _a.generateNewId(),
            webPartFeatureName: `${args.options.webPartTitle} Feature`,
            webPartFeatureDescription: `A feature which activates the Client-Side WebPart named ${args.options.webPartTitle}`,
            webPartAlias: _a.getWebPartAlias(args.options.webPartTitle),
            webPartName: args.options.webPartTitle,
            webPartSafeName: _a.getSafeWebPartName(args.options.webPartTitle),
            webPartDescription: args.options.webPartDescription,
            webPartModule: _a.getSafePackageName(args.options.webPartTitle)
        };
        let tmpDir = undefined;
        let error;
        try {
            if (this.verbose) {
                await logger.log(`Creating temp folder...`);
            }
            tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'cli-spfx'));
            if (this.debug) {
                await logger.log(`Temp folder created at ${tmpDir}`);
            }
            if (this.verbose) {
                await logger.log('Copying files...');
            }
            const src = path.join(__dirname, 'package-generate', 'assets');
            fsUtil.copyRecursiveSync(src, tmpDir, s => _a.replaceTokens(s, tokens));
            const files = fsUtil.readdirR(tmpDir);
            if (this.verbose) {
                await logger.log('Processing files...');
            }
            for (const filePath of files) {
                if (this.debug) {
                    await logger.log(`Processing ${filePath}...`);
                }
                if (!_a.isBinaryFile(filePath)) {
                    if (this.verbose) {
                        await logger.log('Replacing tokens...');
                    }
                    let fileContents = fs.readFileSync(filePath, 'utf-8');
                    if (this.debug) {
                        await logger.log('Before:');
                        await logger.log(fileContents);
                    }
                    fileContents = _a.replaceTokens(fileContents, tokens);
                    if (this.debug) {
                        await logger.log('After:');
                        await logger.log(fileContents);
                    }
                    fs.writeFileSync(filePath, fileContents, { encoding: 'utf-8' });
                }
                else {
                    if (this.verbose) {
                        await logger.log(`Binary file. Skipping replacing tokens in contents`);
                    }
                }
            }
            if (this.verbose) {
                await logger.log('Creating .sppkg file...');
            }
            // we need this to be able to inject mock AdmZip for testing
            /* c8 ignore next 3 */
            if (!this.archive) {
                this.archive = new AdmZip();
            }
            const filesToZip = fsUtil.readdirR(tmpDir);
            for (const f of filesToZip) {
                if (this.debug) {
                    await logger.log(`Adding ${f} to archive...`);
                }
                this.archive.addLocalFile(f, path.relative(tmpDir, path.dirname(f)), path.basename(f));
            }
            if (this.debug) {
                await logger.log('Writing archive...');
            }
            this.archive.writeZip(`${args.options.name}.sppkg`);
        }
        catch (err) {
            error = err.message;
        }
        finally {
            /* eslint-disable no-unsafe-finally */
            try {
                if (tmpDir) {
                    if (this.verbose) {
                        await logger.log(`Deleting temp folder at ${tmpDir}...`);
                    }
                    fs.rmdirSync(tmpDir, { recursive: true });
                }
                if (error) {
                    throw error;
                }
            }
            catch (ex) {
                if (ex === error) {
                    throw ex;
                }
                throw `An error has occurred while removing the temp folder at ${tmpDir}. Please remove it manually.`;
            }
            /* eslint-enable no-unsafe-finally */
        }
    }
    static replaceTokens(s, tokens) {
        return s.replace(/\$([^$]+)\$/g, (substring, token) => {
            if (tokens[token]) {
                return tokens[token];
            }
            else {
                return substring;
            }
        });
    }
    static isBinaryFile(filePath) {
        return filePath.endsWith('.png');
    }
    static getSafePackageName(packageName) {
        return packageName.toLowerCase().replace(/[^a-zA-Z0-9]/g, '-');
    }
    static getSafeWebPartName(webPartName) {
        return webPartName.replace(/ /g, '-');
    }
    static getWebPartAlias(webPartName) {
        return 'AutoWP' + webPartName.replace(/[^a-zA-Z0-9]/g, '').substring(0, 40);
    }
}
_a = SpfxPackageGenerateCommand, _SpfxPackageGenerateCommand_instances = new WeakSet(), _SpfxPackageGenerateCommand_initTelemetry = function _SpfxPackageGenerateCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            allowTenantWideDeployment: args.options.allowTenantWideDeployment === true,
            developerMpnId: typeof args.options.developerMpnId !== 'undefined',
            developerName: typeof args.options.developerName !== 'undefined',
            developerPrivacyUrl: typeof args.options.developerPrivacyUrl !== 'undefined',
            developerTermsOfUseUrl: typeof args.options.developerTermsOfUseUrl !== 'undefined',
            developerWebsiteUrl: typeof args.options.developerWebsiteUrl !== 'undefined',
            enableForTeams: args.options.enableForTeams,
            exposePageContextGlobally: args.options.exposePageContextGlobally === true,
            exposeTeamsContextGlobally: args.options.exposeTeamsContextGlobally === true
        });
    });
}, _SpfxPackageGenerateCommand_initOptions = function _SpfxPackageGenerateCommand_initOptions() {
    this.options.unshift({ option: '-t, --webPartTitle <webPartTitle>' }, { option: '-d, --webPartDescription <webPartDescription>' }, { option: '-n, --name <name>' }, { option: '--html <html>' }, {
        option: '--enableForTeams [enableForTeams]',
        autocomplete: _a.enableForTeamsOptions
    }, { option: '--exposePageContextGlobally' }, { option: '--exposeTeamsContextGlobally' }, { option: '--allowTenantWideDeployment' }, { option: '--developerName [developerName]' }, { option: '--developerPrivacyUrl [developerPrivacyUrl]' }, { option: '--developerTermsOfUseUrl [developerTermsOfUseUrl]' }, { option: '--developerWebsiteUrl [developerWebsiteUrl]' }, { option: '--developerMpnId [developerMpnId]' });
}, _SpfxPackageGenerateCommand_initValidators = function _SpfxPackageGenerateCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.enableForTeams &&
            _a.enableForTeamsOptions.indexOf(args.options.enableForTeams) < 0) {
            return `${args.options.enableForTeams} is not a valid value for enableForTeams. Allowed values are: ${_a.enableForTeamsOptions.join(', ')}`;
        }
        return true;
    });
};
SpfxPackageGenerateCommand.enableForTeamsOptions = ['tab', 'personalApp', 'all'];
SpfxPackageGenerateCommand.generateNewId = () => {
    return v4();
};
export default new SpfxPackageGenerateCommand();
//# sourceMappingURL=package-generate.js.map