var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFileGetCommand_instances, _SpoFileGetCommand_initTelemetry, _SpoFileGetCommand_initOptions, _SpoFileGetCommand_initValidators, _SpoFileGetCommand_initOptionSets, _SpoFileGetCommand_initTypes;
import fs from 'fs';
import path from 'path';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFileGetCommand extends SpoCommand {
    get name() {
        return commands.FILE_GET;
    }
    get description() {
        return 'Gets information about the specified file';
    }
    constructor() {
        super();
        _SpoFileGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFileGetCommand_instances, "m", _SpoFileGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFileGetCommand_instances, "m", _SpoFileGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFileGetCommand_instances, "m", _SpoFileGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFileGetCommand_instances, "m", _SpoFileGetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFileGetCommand_instances, "m", _SpoFileGetCommand_initTypes).call(this);
    }
    getExcludedOptionsWithUrls() {
        return ['url'];
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving file from site ${args.options.webUrl}...`);
        }
        let requestUrl = '';
        let options = '';
        if (args.options.id) {
            requestUrl = `${args.options.webUrl}/_api/web/GetFileById('${formatting.encodeQueryParameter(args.options.id)}')`;
        }
        else if (args.options.url) {
            requestUrl = `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl=@f)`;
        }
        if (args.options.asListItem) {
            options = '?$expand=ListItemAllFields';
        }
        else if (args.options.asFile || args.options.asString) {
            options = '/$value';
        }
        if (args.options.url) {
            if (options.indexOf('?') < 0) {
                options += '?';
            }
            else {
                options += '&';
            }
            const serverRelativePath = urlUtil.getServerRelativePath(args.options.webUrl, args.options.url);
            options += `@f='${formatting.encodeQueryParameter(serverRelativePath)}'`;
        }
        const requestOptions = {
            url: requestUrl + options,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            // Set responseType to arraybuffer, otherwise binary data will be encoded
            // to utf8 and binary data is corrupt
            responseType: args.options.asFile ? 'stream' : 'json'
        };
        try {
            const file = await request.get(requestOptions);
            if (args.options.asFile && args.options.path) {
                // Not possible to use async/await for this promise
                await new Promise((resolve, reject) => {
                    const writer = fs.createWriteStream(args.options.path);
                    file.data.pipe(writer);
                    writer.on('error', err => {
                        reject(err);
                    });
                    writer.on('close', async () => {
                        const filePath = args.options.path;
                        if (this.verbose) {
                            await logger.logToStderr(`File saved to path ${filePath}`);
                        }
                        return resolve();
                    });
                });
            }
            else {
                if (args.options.asString) {
                    await logger.log(file.toString());
                }
                else {
                    const fileProperties = JSON.parse(JSON.stringify(file));
                    if (args.options.withPermissions) {
                        requestOptions.url = `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${file.ServerRelativeUrl}')/ListItemAllFields/RoleAssignments?$expand=Member,RoleDefinitionBindings`;
                        const response = await request.get(requestOptions);
                        response.value.forEach((r) => {
                            r.RoleDefinitionBindings = formatting.setFriendlyPermissions(r.RoleDefinitionBindings);
                        });
                        fileProperties.RoleAssignments = response.value;
                        if (args.options.asListItem) {
                            fileProperties.ListItemAllFields.RoleAssignments = response.value;
                        }
                    }
                    if (args.options.asListItem) {
                        delete fileProperties.ListItemAllFields.ID;
                    }
                    await logger.log(args.options.asListItem ? fileProperties.ListItemAllFields : fileProperties);
                }
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoFileGetCommand_instances = new WeakSet(), _SpoFileGetCommand_initTelemetry = function _SpoFileGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            url: typeof args.options.url !== 'undefined',
            asString: !!args.options.asString,
            asListItem: !!args.options.asListItem,
            asFile: !!args.options.asFile,
            path: typeof args.options.path !== 'undefined',
            withPermissions: !!args.options.withPermissions
        });
    });
}, _SpoFileGetCommand_initOptions = function _SpoFileGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--url [url]'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '--asString'
    }, {
        option: '--asListItem'
    }, {
        option: '--asFile'
    }, {
        option: '-p, --path [path]'
    }, {
        option: '--withPermissions'
    });
}, _SpoFileGetCommand_initValidators = function _SpoFileGetCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.id) {
            if (!validation.isValidGuid(args.options.id)) {
                return `${args.options.id} is not a valid GUID`;
            }
        }
        if (args.options.asFile && !args.options.path) {
            return 'The path should be specified when the --asFile option is used';
        }
        if (args.options.path && !fs.existsSync(path.dirname(args.options.path))) {
            return 'Specified path where to save the file does not exits';
        }
        if (args.options.asFile) {
            if (args.options.asListItem || args.options.asString) {
                return 'Specify to retrieve the file either as file, list item or string but not multiple';
            }
        }
        if (args.options.asListItem) {
            if (args.options.asFile || args.options.asString) {
                return 'Specify to retrieve the file either as file, list item or string but not multiple';
            }
        }
        return true;
    });
}, _SpoFileGetCommand_initOptionSets = function _SpoFileGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'url'] });
}, _SpoFileGetCommand_initTypes = function _SpoFileGetCommand_initTypes() {
    this.types.string.push('webUrl', 'url', 'id', 'path');
    this.types.boolean.push('asString', 'asListItem', 'asFile', 'withPermissions');
};
export default new SpoFileGetCommand();
//# sourceMappingURL=file-get.js.map