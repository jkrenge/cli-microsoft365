var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoPageCopyCommand_instances, _SpoPageCopyCommand_initTelemetry, _SpoPageCopyCommand_initOptions, _SpoPageCopyCommand_initValidators;
import chalk from 'chalk';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoPageCopyCommand extends SpoCommand {
    get name() {
        return `${commands.PAGE_COPY}`;
    }
    get description() {
        return 'Creates a copy of a modern page or template';
    }
    constructor() {
        super();
        _SpoPageCopyCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoPageCopyCommand_instances, "m", _SpoPageCopyCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoPageCopyCommand_instances, "m", _SpoPageCopyCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoPageCopyCommand_instances, "m", _SpoPageCopyCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        let { webUrl } = args.options;
        const { targetUrl, overwrite } = args.options;
        webUrl = this.removeTrailingSlash(webUrl);
        let sourceFullName = args.options.sourceName.toLowerCase();
        const targetPageInfo = this.getTargetSiteUrl(webUrl, targetUrl.toLowerCase());
        let { siteUrl: targetSiteUrl, pageName: targetFullName } = targetPageInfo;
        if (sourceFullName.indexOf('.aspx') < 0) {
            sourceFullName += '.aspx';
        }
        if (targetFullName.indexOf('.aspx') < 0) {
            targetFullName += '.aspx';
        }
        targetSiteUrl = this.removeTrailingSlash(targetSiteUrl);
        if (targetFullName.startsWith('/')) {
            targetFullName = targetFullName.substring(1);
        }
        if (this.verbose) {
            await logger.logToStderr(`Creating page copy...`);
        }
        try {
            let requestOptions = {
                url: `${webUrl}/_api/SP.MoveCopyUtil.CopyFileByPath()`,
                headers: {
                    'accept': 'application/json;odata=nometadata'
                },
                data: {
                    srcPath: { DecodedUrl: `${webUrl}/sitepages/${sourceFullName}` },
                    destPath: { DecodedUrl: `${targetSiteUrl}/sitepages/${targetFullName}` },
                    options: { ResetAuthorAndCreatedOnCopy: true, ShouldBypassSharedLocks: true },
                    overwrite: !!overwrite
                },
                responseType: 'json'
            };
            await request.post(requestOptions);
            requestOptions = {
                url: `${targetSiteUrl}/_api/sitepages/pages/GetByUrl('sitepages/${targetFullName}')`,
                headers: {
                    'accept': 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            const res = await request.get(requestOptions);
            await logger.log(res);
            if (this.verbose) {
                await logger.logToStderr(chalk.green('DONE'));
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getTargetSiteUrl(webUrl, targetFullName) {
        const siteSplit = targetFullName.split('sitepages/');
        if (targetFullName.startsWith("http")) {
            return {
                siteUrl: siteSplit[0],
                pageName: siteSplit[1]
            };
        }
        else {
            return {
                siteUrl: webUrl,
                pageName: siteSplit.length > 1 ? siteSplit[1] : targetFullName
            };
        }
    }
    removeTrailingSlash(value) {
        if (value.endsWith('/')) {
            value = value.substring(0, value.length - 1);
        }
        return value;
    }
}
_SpoPageCopyCommand_instances = new WeakSet(), _SpoPageCopyCommand_initTelemetry = function _SpoPageCopyCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            overwrite: !!args.options.overwrite
        });
    });
}, _SpoPageCopyCommand_initOptions = function _SpoPageCopyCommand_initOptions() {
    this.options.unshift({
        option: '--sourceName <sourceName>'
    }, {
        option: '--targetUrl <targetUrl>'
    }, {
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--overwrite'
    });
}, _SpoPageCopyCommand_initValidators = function _SpoPageCopyCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.webUrl));
};
export default new SpoPageCopyCommand();
//# sourceMappingURL=page-copy.js.map