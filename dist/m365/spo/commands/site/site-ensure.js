var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteEnsureCommand_instances, _SpoSiteEnsureCommand_initOptions, _SpoSiteEnsureCommand_initTypes, _SpoSiteEnsureCommand_initValidators;
import chalk from 'chalk';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { SharingCapabilities } from './SharingCapabilities.js';
class SpoSiteEnsureCommand extends SpoCommand {
    get name() {
        return commands.SITE_ENSURE;
    }
    get description() {
        return 'Ensures that the particular site collection exists and updates its properties if necessary';
    }
    constructor() {
        super();
        _SpoSiteEnsureCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteEnsureCommand_instances, "m", _SpoSiteEnsureCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteEnsureCommand_instances, "m", _SpoSiteEnsureCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _SpoSiteEnsureCommand_instances, "m", _SpoSiteEnsureCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const res = await this.ensureSite(logger, args);
            await logger.log(res);
            if (this.verbose) {
                await logger.logToStderr(chalk.green('DONE'));
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async ensureSite(logger, args) {
        let getWebOutput;
        try {
            getWebOutput = await this.getWeb(args, logger);
        }
        catch (err) {
            if (this.debug) {
                await logger.logToStderr(err);
            }
            if (err.error !== '404 FILE NOT FOUND') {
                throw err;
            }
            if (this.verbose) {
                await logger.logToStderr(`No site found at ${args.options.url}`);
            }
            return this.createSite(args, logger);
        }
        if (this.debug) {
            await logger.logToStderr(getWebOutput);
        }
        if (this.verbose) {
            await logger.logToStderr(`Site found at ${args.options.url}. Checking if site matches conditions...`);
        }
        const web = {
            Configuration: getWebOutput.Configuration,
            WebTemplate: getWebOutput.WebTemplate
        };
        if (args.options.type) {
            // type was specified so we need to check if the existing site matches
            // it. If not, we throw an error and stop
            // Determine the type of site to match
            let expectedWebTemplate;
            switch (args.options.type) {
                case 'TeamSite':
                    expectedWebTemplate = 'GROUP#0';
                    break;
                case 'CommunicationSite':
                    expectedWebTemplate = 'SITEPAGEPUBLISHING#0';
                    break;
                case 'ClassicSite':
                    expectedWebTemplate = args.options.webTemplate || 'STS#0';
                    break;
                default:
                    throw `${args.options.type} is not a valid site type. Allowed types are TeamSite,CommunicationSite,ClassicSite`;
            }
            if (expectedWebTemplate) {
                const currentWebTemplate = `${web.WebTemplate}#${web.Configuration}`;
                if (expectedWebTemplate !== currentWebTemplate) {
                    throw `Expected web template ${expectedWebTemplate} but site found at ${args.options.url} is based on ${currentWebTemplate}`;
                }
            }
        }
        if (this.verbose) {
            await logger.logToStderr(`Site matches conditions. Updating...`);
        }
        return this.updateSite(args, logger);
    }
    async getWeb(args, logger) {
        if (this.verbose) {
            await logger.logToStderr(`Checking if site ${args.options.url} exists...`);
        }
        return await spo.getWeb(args.options.url, logger, this.verbose);
    }
    async createSite(args, logger) {
        if (this.verbose) {
            await logger.logToStderr(`Creating site...`);
        }
        const url = typeof args.options.type === 'undefined' || args.options.type === 'TeamSite' ? undefined : args.options.url;
        return await spo.addSite(args.options.title, logger, this.verbose, args.options.wait, args.options.type, args.options.alias, args.options.description, args.options.owners, args.options.shareByEmailEnabled, args.options.removeDeletedSite, args.options.classification, args.options.isPublic, args.options.lcid, url, args.options.siteDesign, args.options.siteDesignId, args.options.timeZone, args.options.webTemplate, args.options.resourceQuota, args.options.resourceQuotaWarningLevel, args.options.storageQuota, args.options.storageQuotaWarningLevel);
    }
    async updateSite(args, logger) {
        if (this.verbose) {
            await logger.logToStderr(`Updating site...`);
        }
        return await spo.updateSite(args.options.url, logger, this.verbose, args.options.title, args.options.classification, args.options.disableFlows, args.options.isPublic, args.options.owners, args.options.shareByEmailEnabled, args.options.siteDesignId, args.options.sharingCapability);
    }
    /**
     * Maps the base sharingCapability enum to string array so it can
     * more easily be used in validation or descriptions.
     */
    get sharingCapabilities() {
        const result = [];
        for (const sharingCapability in SharingCapabilities) {
            if (typeof SharingCapabilities[sharingCapability] === 'number') {
                result.push(sharingCapability);
            }
        }
        return result;
    }
}
_SpoSiteEnsureCommand_instances = new WeakSet(), _SpoSiteEnsureCommand_initOptions = function _SpoSiteEnsureCommand_initOptions() {
    this.options.unshift({
        option: '-u, --url <url>'
    }, {
        option: '--type [type]',
        autocomplete: ['TeamSite', 'CommunicationSite', 'ClassicSite']
    }, {
        option: '-t, --title <title>'
    }, {
        option: '-a, --alias [alias]'
    }, {
        option: '-z, --timeZone [timeZone]'
    }, {
        option: '-d, --description [description]'
    }, {
        option: '-l, --lcid [lcid]'
    }, {
        option: '--owners [owners]'
    }, {
        option: '--isPublic'
    }, {
        option: '-c, --classification [classification]'
    }, {
        option: '--siteDesign [siteDesign]',
        autocomplete: ['Topic', 'Showcase', 'Blank']
    }, {
        option: '--siteDesignId [siteDesignId]'
    }, {
        option: '--shareByEmailEnabled'
    }, {
        option: '-w, --webTemplate [webTemplate]'
    }, {
        option: '--resourceQuota [resourceQuota]'
    }, {
        option: '--resourceQuotaWarningLevel [resourceQuotaWarningLevel]'
    }, {
        option: '--storageQuota [storageQuota]'
    }, {
        option: '--storageQuotaWarningLevel [storageQuotaWarningLevel]'
    }, {
        option: '--removeDeletedSite'
    }, {
        option: '--disableFlows [disableFlows]',
        autocomplete: ['true', 'false']
    }, {
        option: '--sharingCapability [sharingCapability]',
        autocomplete: this.sharingCapabilities
    }, {
        option: '--wait'
    });
}, _SpoSiteEnsureCommand_initTypes = function _SpoSiteEnsureCommand_initTypes() {
    this.types.boolean.push('disableFlows');
}, _SpoSiteEnsureCommand_initValidators = function _SpoSiteEnsureCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.url));
};
export default new SpoSiteEnsureCommand();
//# sourceMappingURL=site-ensure.js.map