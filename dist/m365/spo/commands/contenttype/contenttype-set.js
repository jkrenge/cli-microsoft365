var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoContentTypeSetCommand_instances, _SpoContentTypeSetCommand_initTelemetry, _SpoContentTypeSetCommand_initOptions, _SpoContentTypeSetCommand_initValidators, _SpoContentTypeSetCommand_initTypes, _SpoContentTypeSetCommand_initOptionSets;
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoContentTypeSetCommand extends SpoCommand {
    get name() {
        return commands.CONTENTTYPE_SET;
    }
    get description() {
        return 'Update an existing content type';
    }
    constructor() {
        super();
        _SpoContentTypeSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoContentTypeSetCommand_instances, "m", _SpoContentTypeSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeSetCommand_instances, "m", _SpoContentTypeSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeSetCommand_instances, "m", _SpoContentTypeSetCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeSetCommand_instances, "m", _SpoContentTypeSetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeSetCommand_instances, "m", _SpoContentTypeSetCommand_initValidators).call(this);
    }
    allowUnknownOptions() {
        return true;
    }
    async commandAction(logger, args) {
        try {
            const contentTypeId = await this.getContentTypeId(logger, args.options);
            const siteId = await spo.getSiteIdBySPApi(args.options.webUrl, logger, this.verbose);
            const webId = await spo.getWebId(args.options.webUrl, logger, this.verbose);
            await this.updateContentType(logger, siteId, webId, contentTypeId, args.options);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getContentTypeId(logger, options) {
        if (options.id) {
            return options.id;
        }
        if (this.verbose) {
            await logger.logToStderr(`Retrieving content type to update...`);
        }
        const requestOptions = {
            url: `${options.webUrl}/_api/Web`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        if (options.listId) {
            requestOptions.url += `/Lists/GetById('${formatting.encodeQueryParameter(options.listId)}')`;
        }
        else if (options.listTitle) {
            requestOptions.url += `/Lists/GetByTitle('${formatting.encodeQueryParameter(options.listTitle)}')`;
        }
        else if (options.listUrl) {
            requestOptions.url += `/GetList('${formatting.encodeQueryParameter(urlUtil.getServerRelativePath(options.webUrl, options.listUrl))}')`;
        }
        requestOptions.url += `/ContentTypes?$filter=Name eq '${formatting.encodeQueryParameter(options.name)}'&$select=Id`;
        const res = await request.get(requestOptions);
        if (res.value.length === 0) {
            throw `The specified content type '${options.name}' does not exist`;
        }
        return res.value[0].Id.StringValue;
    }
    async updateContentType(logger, siteId, webId, contentTypeId, options) {
        if (this.verbose) {
            await logger.logToStderr(`Updating content type...`);
        }
        const requestOptions = {
            url: `${options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
                'Content-Type': 'text/xml'
            },
            data: await this.getCsomCallXmlBody(options, siteId, webId, contentTypeId, logger)
        };
        const res = await request.post(requestOptions);
        const json = JSON.parse(res);
        const response = json[0];
        if (response.ErrorInfo) {
            throw response.ErrorInfo.ErrorMessage;
        }
    }
    async getCsomCallXmlBody(options, siteId, webId, contentTypeId, logger) {
        const payload = this.getRequestPayload(options);
        const list = await this.getListId(options, logger);
        return `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${payload}</Actions><ObjectPaths><Identity Id="9" Name="fc4179a0-e0d7-5000-c38b-bc3506fbab6f|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${siteId}:web:${webId}${list}:contenttype:${formatting.escapeXml(contentTypeId)}" /></ObjectPaths></Request>`;
    }
    getRequestPayload(options) {
        const excludeOptions = [
            'webUrl',
            'id',
            'name',
            'listTitle',
            'listId',
            'listUrl',
            'query',
            'debug',
            'verbose',
            'output',
            'updateChildren'
        ];
        let i = 12;
        const payload = Object.keys(options)
            .filter(key => excludeOptions.indexOf(key) === -1)
            .map(key => {
            return `<SetProperty Id="${i++}" ObjectPathId="9" Name="${key}"><Parameter Type="String">${formatting.escapeXml(options[key])}</Parameter></SetProperty>`;
        });
        if (options.updateChildren) {
            payload.push(`<Method Name="Update" Id="${i++}" ObjectPathId="9"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method>`);
        }
        else {
            payload.push(`<Method Name="Update" Id="${i++}" ObjectPathId="9"><Parameters><Parameter Type="Boolean">false</Parameter></Parameters></Method>`);
        }
        return payload.join('');
    }
    async getListId(options, logger) {
        if (!options.listId && !options.listTitle && !options.listUrl) {
            return '';
        }
        let baseString = ':list:';
        if (options.listId) {
            return baseString + options.listId;
        }
        return baseString + await spo.getListId(options.webUrl, options.listTitle, options.listUrl, logger, this.verbose);
    }
}
_SpoContentTypeSetCommand_instances = new WeakSet(), _SpoContentTypeSetCommand_initTelemetry = function _SpoContentTypeSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listId: typeof args.options.listId !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            updateChildren: args.options.updateChildren
        });
    });
}, _SpoContentTypeSetCommand_initOptions = function _SpoContentTypeSetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-n, --name [name]'
    }, {
        option: '--listTitle [listTitle]'
    }, {
        option: '--listId [listId]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '--updateChildren'
    });
}, _SpoContentTypeSetCommand_initValidators = function _SpoContentTypeSetCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
            return `'${args.options.listId}' is not a valid GUID.`;
        }
        if ((args.options.listId && (args.options.listTitle || args.options.listUrl)) || (args.options.listTitle && args.options.listUrl)) {
            return `Specify either listTitle, listId or listUrl.`;
        }
        if ((args.options.listId || args.options.listTitle || args.options.listUrl) && args.options.updateChildren) {
            return 'It is impossible to pass updateChildren when trying to update a list content type.';
        }
        return true;
    });
}, _SpoContentTypeSetCommand_initTypes = function _SpoContentTypeSetCommand_initTypes() {
    this.types.string.push('id', 'i');
}, _SpoContentTypeSetCommand_initOptionSets = function _SpoContentTypeSetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] });
};
export default new SpoContentTypeSetCommand();
//# sourceMappingURL=contenttype-set.js.map