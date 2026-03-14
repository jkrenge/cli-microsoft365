var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFolderListCommand_instances, _SpoFolderListCommand_initTelemetry, _SpoFolderListCommand_initOptions, _SpoFolderListCommand_initValidators, _SpoFolderListCommand_initTypes;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFolderListCommand extends SpoCommand {
    get name() {
        return commands.FOLDER_LIST;
    }
    get description() {
        return 'Returns all folders under the specified parent folder';
    }
    defaultProperties() {
        return ['Name', 'ServerRelativeUrl'];
    }
    constructor() {
        super();
        _SpoFolderListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFolderListCommand_instances, "m", _SpoFolderListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFolderListCommand_instances, "m", _SpoFolderListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFolderListCommand_instances, "m", _SpoFolderListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFolderListCommand_instances, "m", _SpoFolderListCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving all folders in folder '${args.options.parentFolderUrl}' at site '${args.options.webUrl}'${args.options.recursive ? ' (recursive)' : ''}...`);
        }
        try {
            const fieldProperties = this.formatSelectProperties(args.options.fields);
            const allFiles = await this.getFolders(args.options.parentFolderUrl, fieldProperties, args, logger);
            // Clean ListItemAllFields.ID property from the output if included
            // Reason: It causes a casing conflict with 'Id' when parsing JSON in PowerShell
            if (fieldProperties.selectProperties.some(p => p.toLowerCase().indexOf('listitemallfields') > -1)) {
                allFiles.filter(folder => folder.ListItemAllFields?.ID !== undefined).forEach(folder => delete folder.ListItemAllFields['ID']);
            }
            await logger.log(allFiles);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getFolders(parentFolderUrl, fieldProperties, args, logger, skip = 0) {
        if (this.verbose) {
            const page = Math.ceil(skip / SpoFolderListCommand.pageSize) + 1;
            await logger.logToStderr(`Retrieving folders in folder '${parentFolderUrl}'${page > 1 ? ', page ' + page : ''}...`);
        }
        const allFolders = [];
        const serverRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, parentFolderUrl);
        const requestUrl = `${args.options.webUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}')/Folders`;
        const queryParams = [`$skip=${skip}`, `$top=${SpoFolderListCommand.pageSize}`];
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
        for (const folder of response.value) {
            allFolders.push(folder);
            if (args.options.recursive) {
                const subFolders = await this.getFolders(folder.ServerRelativeUrl, fieldProperties, args, logger);
                subFolders.forEach(subFolder => allFolders.push(subFolder));
            }
        }
        if (response.value.length === SpoFolderListCommand.pageSize) {
            const folders = await this.getFolders(parentFolderUrl, fieldProperties, args, logger, skip + SpoFolderListCommand.pageSize);
            folders.forEach(folder => allFolders.push(folder));
        }
        return allFolders;
    }
    formatSelectProperties(fields) {
        const selectProperties = [];
        const expandProperties = [];
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
_SpoFolderListCommand_instances = new WeakSet(), _SpoFolderListCommand_initTelemetry = function _SpoFolderListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            recursive: !!args.options.recursive,
            fields: typeof args.options.fields !== 'undefined',
            filter: typeof args.options.filter !== 'undefined'
        });
    });
}, _SpoFolderListCommand_initOptions = function _SpoFolderListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-p, --parentFolderUrl <parentFolderUrl>'
    }, {
        option: '--fields [fields]'
    }, {
        option: '--filter [filter]'
    }, {
        option: '-r, --recursive'
    });
}, _SpoFolderListCommand_initValidators = function _SpoFolderListCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.webUrl));
}, _SpoFolderListCommand_initTypes = function _SpoFolderListCommand_initTypes() {
    this.types.string.push('webUrl', 'parentFolderUrl', 'fields', 'filter');
    this.types.boolean.push('recursive');
};
SpoFolderListCommand.pageSize = 5000;
export default new SpoFolderListCommand();
//# sourceMappingURL=folder-list.js.map