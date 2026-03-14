var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFolderAddCommand_instances, _SpoFolderAddCommand_initTelemetry, _SpoFolderAddCommand_initOptions, _SpoFolderAddCommand_initValidators, _SpoFolderAddCommand_initTypes;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { FolderColorValues } from './FolderColor.js';
class SpoFolderAddCommand extends SpoCommand {
    get name() {
        return commands.FOLDER_ADD;
    }
    get description() {
        return 'Creates a folder within a parent folder';
    }
    constructor() {
        super();
        _SpoFolderAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFolderAddCommand_instances, "m", _SpoFolderAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFolderAddCommand_instances, "m", _SpoFolderAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFolderAddCommand_instances, "m", _SpoFolderAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFolderAddCommand_instances, "m", _SpoFolderAddCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            const parentFolderServerRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.parentFolderUrl);
            const serverRelativeUrl = `${parentFolderServerRelativeUrl}/${args.options.name}`;
            if (args.options.ensureParentFolders) {
                await this.ensureParentFolderPath(args.options, parentFolderServerRelativeUrl, logger);
            }
            const folder = await this.addFolder(serverRelativeUrl, args.options, logger);
            await logger.log(folder);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async ensureParentFolderPath(options, parentFolderPath, logger) {
        if (this.verbose) {
            await logger.logToStderr(`Ensuring parent folders exist...`);
        }
        const parsedUrl = new URL(options.webUrl);
        const absoluteFolderUrl = `${parsedUrl.origin}${parentFolderPath}`;
        const relativeFolderPath = absoluteFolderUrl.replace(options.webUrl, '');
        const parentFolders = relativeFolderPath.split('/').filter(folder => folder !== '');
        for (let i = 1; i < parentFolders.length; i++) {
            const currentFolderPath = parentFolders.slice(0, i + 1).join('/');
            if (this.verbose) {
                await logger.logToStderr(`Checking if folder '${currentFolderPath}' exists...`);
            }
            const folderExists = await this.getFolderExists(options.webUrl, currentFolderPath);
            if (!folderExists) {
                await this.addFolder(currentFolderPath, options, logger);
            }
        }
    }
    async getFolderExists(webUrl, folderServerRelativeUrl) {
        const requestUrl = `${webUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(folderServerRelativeUrl)}')?$select=Exists`;
        const requestOptions = {
            url: requestUrl,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        return response.Exists;
    }
    async addFolder(serverRelativeUrl, options, logger) {
        if (this.verbose) {
            const folderName = serverRelativeUrl.split('/').pop();
            const folderLocation = serverRelativeUrl.split('/').slice(0, -1).join('/');
            await logger.logToStderr(`Adding folder with name '${folderName}' at location '${folderLocation}'...`);
        }
        const requestOptions = {
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        if (options.color === undefined) {
            requestOptions.url = `${options.webUrl}/_api/web/folders/addUsingPath(decodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}')`;
        }
        else {
            requestOptions.url = `${options.webUrl}/_api/foldercoloring/createfolder(DecodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}', overwrite=false)`;
            requestOptions.data = {
                coloringInformation: {
                    ColorHex: FolderColorValues[options.color] || options.color
                }
            };
        }
        const response = await request.post(requestOptions);
        return response;
    }
}
_SpoFolderAddCommand_instances = new WeakSet(), _SpoFolderAddCommand_initTelemetry = function _SpoFolderAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            color: typeof args.options.color !== 'undefined',
            ensureParentFolders: !!args.options.ensureParentFolders
        });
    });
}, _SpoFolderAddCommand_initOptions = function _SpoFolderAddCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-p, --parentFolderUrl <parentFolderUrl>'
    }, {
        option: '-n, --name <name>'
    }, {
        option: '--color [color]',
        autocomplete: Object.keys(FolderColorValues)
    }, {
        option: '--ensureParentFolders [ensureParentFolders]'
    });
}, _SpoFolderAddCommand_initValidators = function _SpoFolderAddCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.color && !Object.entries(FolderColorValues).flat().includes(args.options.color)) {
            return `'${args.options.color}' is not a valid value for option 'color'. Allowed values are ${Object.keys(FolderColorValues).join(', ')}, ${Object.values(FolderColorValues).join(', ')}.`;
        }
        return true;
    });
}, _SpoFolderAddCommand_initTypes = function _SpoFolderAddCommand_initTypes() {
    this.types.string.push('webUrl', 'parentFolderUrl', 'name', 'color');
    this.types.boolean.push('ensureParentFolders');
};
export default new SpoFolderAddCommand();
//# sourceMappingURL=folder-add.js.map