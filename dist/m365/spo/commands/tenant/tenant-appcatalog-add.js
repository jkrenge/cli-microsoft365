var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoTenantAppCatalogAddCommand_instances, _SpoTenantAppCatalogAddCommand_initTelemetry, _SpoTenantAppCatalogAddCommand_initOptions, _SpoTenantAppCatalogAddCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import { CommandError } from '../../../../Command.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import spoSiteAddCommand from '../site/site-add.js';
import spoSiteGetCommand from '../site/site-get.js';
import spoSiteRemoveCommand from '../site/site-remove.js';
import spoTenantAppCatalogUrlGetCommand from './tenant-appcatalogurl-get.js';
class SpoTenantAppCatalogAddCommand extends SpoCommand {
    get name() {
        return commands.TENANT_APPCATALOG_ADD;
    }
    get description() {
        return 'Creates new tenant app catalog site';
    }
    constructor() {
        super();
        _SpoTenantAppCatalogAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoTenantAppCatalogAddCommand_instances, "m", _SpoTenantAppCatalogAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoTenantAppCatalogAddCommand_instances, "m", _SpoTenantAppCatalogAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoTenantAppCatalogAddCommand_instances, "m", _SpoTenantAppCatalogAddCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr('Checking for existing app catalog URL...');
        }
        const spoTenantAppCatalogUrlGetCommandOutput = await cli.executeCommandWithOutput(spoTenantAppCatalogUrlGetCommand, { options: { output: 'text', _: [] } });
        const appCatalogUrl = spoTenantAppCatalogUrlGetCommandOutput.stdout;
        if (!appCatalogUrl) {
            if (this.verbose) {
                await logger.logToStderr('No app catalog URL found');
            }
        }
        else {
            if (this.verbose) {
                await logger.logToStderr(`Found app catalog URL ${appCatalogUrl}`);
            }
            //Using JSON.parse
            await this.ensureNoExistingSite(appCatalogUrl, args.options.force, logger);
        }
        await this.ensureNoExistingSite(args.options.url, args.options.force, logger);
        await this.createAppCatalog(args.options, logger);
    }
    async ensureNoExistingSite(url, force, logger) {
        if (this.verbose) {
            await logger.logToStderr(`Checking if site ${url} exists...`);
        }
        const siteGetOptions = {
            options: {
                url: url,
                verbose: this.verbose,
                debug: this.debug,
                _: []
            }
        };
        try {
            await cli.executeCommandWithOutput(spoSiteGetCommand, siteGetOptions);
            if (this.verbose) {
                await logger.logToStderr(`Found site ${url}`);
            }
            if (!force) {
                throw new CommandError(`Another site exists at ${url}`);
            }
            if (this.verbose) {
                await logger.logToStderr(`Deleting site ${url}...`);
            }
            const siteRemoveOptions = {
                url: url,
                skipRecycleBin: true,
                wait: true,
                force: true,
                verbose: this.verbose,
                debug: this.debug
            };
            await cli.executeCommand(spoSiteRemoveCommand, { options: { ...siteRemoveOptions, _: [] } });
        }
        catch (err) {
            if (err.error?.message !== 'File not Found' && err.error?.message !== '404 FILE NOT FOUND') {
                throw err.error || err;
            }
            if (this.verbose) {
                await logger.logToStderr(`No site found at ${url}`);
            }
            // Site not found. Continue
        }
    }
    async createAppCatalog(options, logger) {
        if (this.verbose) {
            await logger.logToStderr(`Creating app catalog at ${options.url}...`);
        }
        const siteAddOptions = {
            webTemplate: 'APPCATALOG#0',
            title: 'App catalog',
            type: 'ClassicSite',
            url: options.url,
            timeZone: options.timeZone,
            owners: options.owner,
            wait: options.wait,
            verbose: this.verbose,
            debug: this.debug,
            removeDeletedSite: false
        };
        return cli.executeCommand(spoSiteAddCommand, { options: { ...siteAddOptions, _: [] } });
    }
}
_SpoTenantAppCatalogAddCommand_instances = new WeakSet(), _SpoTenantAppCatalogAddCommand_initTelemetry = function _SpoTenantAppCatalogAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            wait: args.options.wait || false,
            force: args.options.force || false
        });
    });
}, _SpoTenantAppCatalogAddCommand_initOptions = function _SpoTenantAppCatalogAddCommand_initOptions() {
    this.options.unshift({
        option: '-u, --url <url>'
    }, {
        option: '--owner <owner>'
    }, {
        option: '-z, --timeZone <timeZone>'
    }, {
        option: '--wait'
    }, {
        option: '--force'
    });
}, _SpoTenantAppCatalogAddCommand_initValidators = function _SpoTenantAppCatalogAddCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.url);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (typeof args.options.timeZone !== 'number') {
            return `${args.options.timeZone} is not a number`;
        }
        return true;
    });
};
export default new SpoTenantAppCatalogAddCommand();
//# sourceMappingURL=tenant-appcatalog-add.js.map