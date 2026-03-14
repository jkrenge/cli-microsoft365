var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoAppTeamsPackageDownloadCommand_instances, _SpoAppTeamsPackageDownloadCommand_initTelemetry, _SpoAppTeamsPackageDownloadCommand_initOptions, _SpoAppTeamsPackageDownloadCommand_initValidators;
import fs from 'fs';
import path from 'path';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import { SpoAppBaseCommand } from './SpoAppBaseCommand.js';
class SpoAppTeamsPackageDownloadCommand extends SpoAppBaseCommand {
    get name() {
        return commands.APP_TEAMSPACKAGE_DOWNLOAD;
    }
    get description() {
        return 'Downloads Teams app package for an SPFx solution deployed to tenant app catalog';
    }
    constructor() {
        super();
        _SpoAppTeamsPackageDownloadCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoAppTeamsPackageDownloadCommand_instances, "m", _SpoAppTeamsPackageDownloadCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoAppTeamsPackageDownloadCommand_instances, "m", _SpoAppTeamsPackageDownloadCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoAppTeamsPackageDownloadCommand_instances, "m", _SpoAppTeamsPackageDownloadCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            this.appCatalogUrl = args.options.appCatalogUrl;
            const appInfo = {
                id: args.options.appItemId ?? undefined,
                packageFileName: args.options.fileName ?? undefined
            };
            if (this.debug) {
                await logger.logToStderr(`appInfo: ${JSON.stringify(appInfo)}`);
            }
            await this.ensureAppInfo(logger, args, appInfo);
            if (this.debug) {
                await logger.logToStderr(`ensureAppInfo: ${JSON.stringify(appInfo)}`);
            }
            await this.loadAppCatalogUrl(logger, args);
            const requestOptions = {
                url: `${this.appCatalogUrl}/_api/web/tenantappcatalog/downloadteamssolution(${appInfo.id})/$value`,
                headers: {
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'stream'
            };
            const file = await request.get(requestOptions);
            // Not possible to use async/await for this promise
            await new Promise((resolve, reject) => {
                const writer = fs.createWriteStream(appInfo.packageFileName);
                file.data.pipe(writer);
                writer.on('error', err => {
                    return reject(err);
                });
                writer.on('close', async () => {
                    const fileName = appInfo.packageFileName;
                    if (this.verbose) {
                        await logger.logToStderr(`Package saved to ${fileName}`);
                    }
                    return resolve();
                });
            });
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async ensureAppInfo(logger, args, appInfo) {
        if (appInfo.id && appInfo.packageFileName) {
            return;
        }
        if (args.options.appName && !appInfo.packageFileName) {
            appInfo.packageFileName = this.getPackageNameFromFileName(args.options.appName);
        }
        await this.loadAppCatalogUrl(logger, args);
        const appCatalogListName = 'AppCatalog';
        const serverRelativeAppCatalogListUrl = `${urlUtil.getServerRelativeSiteUrl(this.appCatalogUrl)}/${appCatalogListName}`;
        let url = `${this.appCatalogUrl}/_api/web/`;
        if (args.options.appItemUniqueId) {
            url += `GetList('${serverRelativeAppCatalogListUrl}')/GetItemByUniqueId('${args.options.appItemUniqueId}')?$expand=File&$select=Id,File/Name`;
        }
        else if (args.options.appItemId) {
            url += `GetList('${serverRelativeAppCatalogListUrl}')/GetItemById(${args.options.appItemId})?$expand=File&$select=File/Name`;
        }
        else if (args.options.appName) {
            url += `GetFolderByServerRelativePath(DecodedUrl='${appCatalogListName}')/files('${formatting.encodeQueryParameter(args.options.appName)}')/ListItemAllFields?$select=Id`;
        }
        const requestOptions = {
            url,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const res = await request.get(requestOptions);
        if (args.options.appItemUniqueId) {
            appInfo.id = parseInt(res.Id);
            if (!appInfo.packageFileName) {
                appInfo.packageFileName = this.getPackageNameFromFileName(res.File.Name);
            }
            return;
        }
        if (args.options.appItemId) {
            if (!appInfo.packageFileName) {
                appInfo.packageFileName = this.getPackageNameFromFileName(res.File.Name);
            }
            return;
        }
        appInfo.id = parseInt(res.Id);
    }
    getPackageNameFromFileName(fileName) {
        return `${path.basename(fileName, path.extname(fileName))}.zip`;
    }
    async loadAppCatalogUrl(logger, args) {
        if (this.appCatalogUrl) {
            return;
        }
        const spoUrl = await spo.getSpoUrl(logger, this.debug);
        const appCatalogUrl = await this.getAppCatalogSiteUrl(logger, spoUrl, args);
        this.appCatalogUrl = appCatalogUrl;
    }
}
_SpoAppTeamsPackageDownloadCommand_instances = new WeakSet(), _SpoAppTeamsPackageDownloadCommand_initTelemetry = function _SpoAppTeamsPackageDownloadCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            appCatalogUrl: typeof args.options.appCatalogUrl !== 'undefined',
            appItemUniqueId: typeof args.options.appItemUniqueId !== 'undefined',
            appItemId: typeof args.options.appItemId !== 'undefined',
            appName: typeof args.options.appName !== 'undefined',
            fileName: typeof args.options.fileName !== 'undefined'
        });
    });
}, _SpoAppTeamsPackageDownloadCommand_initOptions = function _SpoAppTeamsPackageDownloadCommand_initOptions() {
    this.options.unshift({ option: '--appItemId [appItemId]' }, { option: '--appItemUniqueId [appItemUniqueId]' }, { option: '--appName [appName]' }, { option: '--fileName [fileName]' }, { option: '-u, --appCatalogUrl [appCatalogUrl]' });
}, _SpoAppTeamsPackageDownloadCommand_initValidators = function _SpoAppTeamsPackageDownloadCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!args.options.appItemUniqueId &&
            !args.options.appItemId &&
            !args.options.appName) {
            return `Specify appItemUniqueId, appItemId or appName`;
        }
        if ((args.options.appItemUniqueId && args.options.appItemId) ||
            (args.options.appItemUniqueId && args.options.appName) ||
            (args.options.appItemId && args.options.appName)) {
            return `Specify appItemUniqueId, appItemId or appName but not multiple`;
        }
        if (args.options.appItemUniqueId &&
            !validation.isValidGuid(args.options.appItemUniqueId)) {
            return `${args.options.appItemUniqueId} is not a valid GUID`;
        }
        if (args.options.appItemId &&
            isNaN(args.options.appItemId)) {
            return `${args.options.appItemId} is not a number`;
        }
        if (args.options.fileName &&
            fs.existsSync(args.options.fileName)) {
            return `File ${args.options.fileName} already exists`;
        }
        if (args.options.appCatalogUrl) {
            return validation.isValidSharePointUrl(args.options.appCatalogUrl);
        }
        return true;
    });
};
export default new SpoAppTeamsPackageDownloadCommand();
//# sourceMappingURL=app-teamspackage-download.js.map