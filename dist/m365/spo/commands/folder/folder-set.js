var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFolderSetCommand_instances, _SpoFolderSetCommand_initTelemetry, _SpoFolderSetCommand_initOptions, _SpoFolderSetCommand_initValidators, _SpoFolderSetCommand_initTypes;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { FolderColorValues } from './FolderColor.js';
class SpoFolderSetCommand extends SpoCommand {
    get name() {
        return commands.FOLDER_SET;
    }
    get description() {
        return 'Updates a folder';
    }
    constructor() {
        super();
        _SpoFolderSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFolderSetCommand_instances, "m", _SpoFolderSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFolderSetCommand_instances, "m", _SpoFolderSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFolderSetCommand_instances, "m", _SpoFolderSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFolderSetCommand_instances, "m", _SpoFolderSetCommand_initTypes).call(this);
    }
    getExcludedOptionsWithUrls() {
        return ['url'];
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Updating folder '${args.options.name}'...`);
            }
            const serverRelativePath = urlUtil.getServerRelativePath(args.options.webUrl, args.options.url);
            if (!args.options.color) {
                const requestOptions = {
                    url: `${args.options.webUrl}/_api/Web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')/ListItemAllFields`,
                    headers: {
                        accept: 'application/json;odata=nometadata',
                        'if-match': '*'
                    },
                    data: {
                        FileLeafRef: args.options.name,
                        Title: args.options.name
                    },
                    responseType: 'json'
                };
                const response = await request.patch(requestOptions);
                if (response && response['odata.null'] === true) {
                    throw 'Folder not found.';
                }
            }
            else {
                const requestOptions = {
                    url: `${args.options.webUrl}/_api/foldercoloring/${args.options.name ? 'renamefolder' : 'stampcolor'}(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')`,
                    headers: {
                        accept: 'application/json;odata=nometadata'
                    },
                    responseType: 'json',
                    data: {
                        coloringInformation: {
                            ColorHex: FolderColorValues[args.options.color] || args.options.color
                        },
                        newName: args.options.name
                    }
                };
                await request.post(requestOptions);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoFolderSetCommand_instances = new WeakSet(), _SpoFolderSetCommand_initTelemetry = function _SpoFolderSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            name: typeof args.options.name !== 'undefined',
            color: typeof args.options.color !== 'undefined'
        });
    });
}, _SpoFolderSetCommand_initOptions = function _SpoFolderSetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--url <url>'
    }, {
        option: '-n, --name [name]'
    }, {
        option: '--color [color]',
        autocomplete: Object.keys(FolderColorValues)
    });
}, _SpoFolderSetCommand_initValidators = function _SpoFolderSetCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.color === undefined && args.options.name === undefined) {
            return `Specify at least one of the options: name or color.`;
        }
        if (args.options.color && !Object.entries(FolderColorValues).flat().includes(args.options.color)) {
            return `'${args.options.color}' is not a valid value for option 'color'. Allowed values are ${Object.keys(FolderColorValues).join(', ')}, ${Object.values(FolderColorValues).join(', ')}.`;
        }
        return true;
    });
}, _SpoFolderSetCommand_initTypes = function _SpoFolderSetCommand_initTypes() {
    this.types.string.push('webUrl', 'url', 'name', 'color');
};
export default new SpoFolderSetCommand();
//# sourceMappingURL=folder-set.js.map