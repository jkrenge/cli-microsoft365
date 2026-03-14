var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteRenameCommand_instances, _SpoSiteRenameCommand_initTelemetry, _SpoSiteRenameCommand_initOptions, _SpoSiteRenameCommand_initValidators;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { timersUtil } from '../../../../utils/timersUtil.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteRenameCommand extends SpoCommand {
    get name() {
        return commands.SITE_RENAME;
    }
    get description() {
        return 'Renames the URL and title of a site collection';
    }
    constructor() {
        super();
        _SpoSiteRenameCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteRenameCommand_instances, "m", _SpoSiteRenameCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteRenameCommand_instances, "m", _SpoSiteRenameCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteRenameCommand_instances, "m", _SpoSiteRenameCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const options = args.options;
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            const reqDigest = await spo.getRequestDigest(spoAdminUrl);
            this.context = reqDigest;
            if (this.verbose) {
                await logger.logToStderr(`Scheduling rename job...`);
            }
            let optionsBitmask = 0;
            if (options.suppressMarketplaceAppCheck) {
                optionsBitmask = optionsBitmask | 8;
            }
            if (options.suppressWorkflow2013Check) {
                optionsBitmask = optionsBitmask | 16;
            }
            const requestOptions = {
                url: `${spoAdminUrl}/_api/SiteRenameJobs?api-version=1.4.7`,
                headers: {
                    'X-RequestDigest': this.context.FormDigestValue,
                    'Content-Type': 'application/json'
                },
                responseType: 'json',
                data: {
                    SourceSiteUrl: options.url,
                    TargetSiteUrl: options.newUrl,
                    TargetSiteTitle: options.newTitle || null,
                    Option: optionsBitmask,
                    Reserve: null,
                    SkipGestures: null,
                    OperationId: '00000000-0000-0000-0000-000000000000'
                }
            };
            const res = await request.post(requestOptions);
            if (options.verbose) {
                await logger.logToStderr(res);
            }
            this.operationData = res;
            if (this.operationData.JobState && this.operationData.JobState === "Error") {
                throw this.operationData.ErrorDescription;
            }
            const isComplete = this.operationData.JobState === "Success";
            if (options.wait && !isComplete) {
                await this.waitForRenameCompletion(this, true, spoAdminUrl, options.url, 0);
            }
            await logger.log(this.operationData);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async waitForRenameCompletion(command, isVerbose, spoAdminUrl, siteUrl, iteration) {
        iteration++;
        const requestOptions = {
            url: `${spoAdminUrl}/_api/SiteRenameJobs/GetJobsBySiteUrl(url='${formatting.encodeQueryParameter(siteUrl)}')?api-version=1.4.7`,
            headers: {
                'X-AttemptNumber': iteration.toString()
            },
            responseType: 'json'
        };
        const res = await request.get(requestOptions);
        this.operationData = res.value[0];
        if (this.operationData.ErrorDescription) {
            throw this.operationData.ErrorDescription;
        }
        if (this.operationData.JobState === "Success") {
            return;
        }
        await timersUtil.setTimeout(SpoSiteRenameCommand.checkIntervalInMs);
        await command.waitForRenameCompletion(command, isVerbose, spoAdminUrl, siteUrl, iteration);
    }
}
_SpoSiteRenameCommand_instances = new WeakSet(), _SpoSiteRenameCommand_initTelemetry = function _SpoSiteRenameCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            newTitle: typeof args.options.newTitle !== 'undefined',
            suppressMarketplaceAppCheck: !!args.options.suppressMarketplaceAppCheck,
            suppressWorkflow2013Check: !!args.options.suppressWorkflow2013Check,
            wait: !!args.options.wait
        });
    });
}, _SpoSiteRenameCommand_initOptions = function _SpoSiteRenameCommand_initOptions() {
    this.options.unshift({
        option: '-u, --url <url>'
    }, {
        option: '--newUrl <newUrl>'
    }, {
        option: '--newTitle [newTitle]'
    }, {
        option: '--suppressMarketplaceAppCheck'
    }, {
        option: '--suppressWorkflow2013Check'
    }, {
        option: '--wait'
    });
}, _SpoSiteRenameCommand_initValidators = function _SpoSiteRenameCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.url.toLowerCase() === args.options.newUrl.toLowerCase()) {
            return 'The new URL cannot be the same as the target URL.';
        }
        return true;
    });
};
SpoSiteRenameCommand.checkIntervalInMs = 5000;
export default new SpoSiteRenameCommand();
//# sourceMappingURL=site-rename.js.map