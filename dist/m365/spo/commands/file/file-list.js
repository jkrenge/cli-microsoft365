var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFileListCommand_instances, _SpoFileListCommand_initTelemetry, _SpoFileListCommand_initOptions, _SpoFileListCommand_initValidators, _SpoFileListCommand_initTypes;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFileListCommand extends SpoCommand {
    get name() {
        return commands.FILE_LIST;
    }
    get description() {
        return 'Lists all available files in the specified folder and site';
    }
    constructor() {
        super();
        _SpoFileListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFileListCommand_instances, "m", _SpoFileListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFileListCommand_instances, "m", _SpoFileListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFileListCommand_instances, "m", _SpoFileListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFileListCommand_instances, "m", _SpoFileListCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving all files in folder '${args.options.folderUrl}' at site '${args.options.webUrl}'${args.options.recursive ? ' (recursive)' : ''}...`);
        }
        try {
            const fieldProperties = this.formatSelectProperties(args.options.fields, args.options.output);
            const allFiles = [];
            const allFolders = args.options.recursive
                ? [...await this.getFolders(args.options.folderUrl, args, logger), args.options.folderUrl]
                : [args.options.folderUrl];
            for (const folder of allFolders) {
                const files = await this.getFiles(folder, fieldProperties, args, logger);
                files.forEach((file) => allFiles.push(file));
            }
            // Clean ListItemAllFields.ID property from the output if included
            // Reason: It causes a casing conflict with 'Id' when parsing JSON in PowerShell
            if (fieldProperties.selectProperties.some(p => p.toLowerCase().indexOf('listitemallfields') > -1)) {
                allFiles.filter(file => file.ListItemAllFields?.ID !== undefined).forEach(file => delete file.ListItemAllFields['ID']);
            }
            await logger.log(allFiles);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getFiles(folderUrl, fieldProperties, args, logger, skip = 0) {
        if (this.verbose) {
            const page = Math.ceil(skip / SpoFileListCommand.pageSize) + 1;
            await logger.logToStderr(`Retrieving files in folder '${folderUrl}'${page > 1 ? ', page ' + page : ''}...`);
        }
        const allFiles = [];
        const serverRelativePath = urlUtil.getServerRelativePath(args.options.webUrl, folderUrl);
        const requestUrl = `${args.options.webUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')/Files`;
        const queryParams = [`$skip=${skip}`, `$top=${SpoFileListCommand.pageSize}`];
        if (fieldProperties.expandProperties.length > 0) {
            queryParams.push(`$expand=${fieldProperties.expandProperties.join(',')}`);
        }
        if (fieldProperties.selectProperties.length > 0) {
            queryParams.push(`$select=${fieldProperties.selectProperties.join(',')}`);
        }
        if (args.options.filter) {
            queryParams.push(`$filter=${args.options.filter}`);
        }
        const requestOptions = {
            url: `${requestUrl}?${queryParams.join('&')}`,
            method: 'GET',
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        response.value.forEach(file => allFiles.push(file));
        if (response.value.length === SpoFileListCommand.pageSize) {
            const files = await this.getFiles(folderUrl, fieldProperties, args, logger, skip + SpoFileListCommand.pageSize);
            files.forEach(file => allFiles.push(file));
        }
        return allFiles;
    }
    async getFolders(folderUrl, args, logger, skip = 0) {
        if (this.verbose) {
            const page = Math.ceil(skip / SpoFileListCommand.pageSize) + 1;
            await logger.logToStderr(`Retrieving folders in folder '${folderUrl}'${page > 1 ? ', page ' + page : ''}...`);
        }
        const allFolders = [];
        const serverRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, folderUrl);
        const requestUrl = `${args.options.webUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}')/Folders`;
        const requestOptions = {
            url: `${requestUrl}?$skip=${skip}&$top=${SpoFileListCommand.pageSize}&$select=ServerRelativeUrl`,
            method: 'GET',
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        for (const folder of response.value) {
            allFolders.push(folder.ServerRelativeUrl);
            const subfolders = await this.getFolders(folder.ServerRelativeUrl, args, logger);
            subfolders.forEach(folder => allFolders.push(folder));
        }
        if (response.value.length === SpoFileListCommand.pageSize) {
            const folders = await this.getFolders(folderUrl, args, logger, skip + SpoFileListCommand.pageSize);
            folders.forEach(folder => allFolders.push(folder));
        }
        return allFolders;
    }
    formatSelectProperties(fields, output) {
        let selectProperties = [];
        const expandProperties = [];
        if (output === 'text' && !fields) {
            selectProperties = ['UniqueId', 'Name', 'ServerRelativeUrl'];
        }
        if (fields) {
            fields.split(',').forEach((field) => {
                const subparts = field.trim().split('/');
                if (subparts.length > 1) {
                    expandProperties.push(subparts[0]);
                }
                selectProperties.push(field.trim());
            });
        }
        return {
            selectProperties: [...new Set(selectProperties)],
            expandProperties: [...new Set(expandProperties)]
        };
    }
}
_SpoFileListCommand_instances = new WeakSet(), _SpoFileListCommand_initTelemetry = function _SpoFileListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            recursive: args.options.recursive,
            fields: typeof args.options.fields !== 'undefined',
            filter: typeof args.options.filter !== 'undefined'
        });
    });
}, _SpoFileListCommand_initOptions = function _SpoFileListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-f, --folderUrl <folderUrl>'
    }, {
        option: '--fields [fields]'
    }, {
        option: '--filter [filter]'
    }, {
        option: '-r, --recursive'
    });
}, _SpoFileListCommand_initValidators = function _SpoFileListCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.webUrl));
}, _SpoFileListCommand_initTypes = function _SpoFileListCommand_initTypes() {
    this.types.string.push('webUrl', 'folderUrl', 'fields', 'filter');
    this.types.boolean.push('recursive');
};
SpoFileListCommand.pageSize = 5000;
export default new SpoFileListCommand();
//# sourceMappingURL=file-list.js.map