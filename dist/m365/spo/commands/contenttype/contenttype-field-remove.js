var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoContentTypeFieldRemoveCommand_instances, _SpoContentTypeFieldRemoveCommand_initTelemetry, _SpoContentTypeFieldRemoveCommand_initOptions, _SpoContentTypeFieldRemoveCommand_initValidators, _SpoContentTypeFieldRemoveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoContentTypeFieldRemoveCommand extends SpoCommand {
    get name() {
        return commands.CONTENTTYPE_FIELD_REMOVE;
    }
    get description() {
        return 'Removes a column from a site- or list content type';
    }
    constructor() {
        super();
        _SpoContentTypeFieldRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoContentTypeFieldRemoveCommand_instances, "m", _SpoContentTypeFieldRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeFieldRemoveCommand_instances, "m", _SpoContentTypeFieldRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeFieldRemoveCommand_instances, "m", _SpoContentTypeFieldRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeFieldRemoveCommand_instances, "m", _SpoContentTypeFieldRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const removeFieldLink = async () => {
            try {
                if (this.debug) {
                    await logger.logToStderr(`Get SiteId required by ProcessQuery endpoint.`);
                }
                const siteId = await spo.getSiteIdBySPApi(args.options.webUrl, logger, this.verbose);
                if (this.debug) {
                    await logger.logToStderr(`SiteId: ${siteId}`);
                    await logger.logToStderr(`Get WebId required by ProcessQuery endpoint.`);
                }
                const webId = await spo.getWebId(args.options.webUrl, logger, this.verbose);
                if (this.debug) {
                    await logger.logToStderr(`WebId: ${webId}`);
                }
                let listId = undefined;
                if (args.options.listId) {
                    listId = args.options.listId;
                }
                if (args.options.listTitle || args.options.listUrl) {
                    listId = await spo.getListId(args.options.webUrl, args.options.listTitle, args.options.listUrl, logger, this.verbose);
                }
                if (this.debug) {
                    await logger.logToStderr(`ListId: ${listId}`);
                }
                const reqDigest = await spo.getRequestDigest(args.options.webUrl);
                const requestDigest = reqDigest.FormDigestValue;
                const updateChildContentTypes = args.options.listTitle || args.options.listId || args.options.listUrl ? false : args.options.updateChildContentTypes === true;
                if (this.debug) {
                    const additionalLog = args.options.listTitle ? `; ListTitle='${args.options.listTitle}'` : args.options.listId ? `; ListId='${args.options.listId}'` : args.options.listUrl ? `; ListUrl='${args.options.listUrl}'` : ` ; UpdateChildContentTypes='${updateChildContentTypes}`;
                    await logger.logToStderr(`Remove FieldLink from ContentType. Id='${args.options.id}' ; ContentTypeId='${args.options.contentTypeId}' ${additionalLog}`);
                    await logger.logToStderr(`Execute ProcessQuery.`);
                }
                let requestBody = '';
                if (listId) {
                    requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName=".NET Library" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="18" ObjectPathId="17" /><ObjectPath Id="20" ObjectPathId="19" /><Method Name="DeleteObject" Id="21" ObjectPathId="19" /><Method Name="Update" Id="22" ObjectPathId="15"><Parameters><Parameter Type="Boolean">${updateChildContentTypes}</Parameter></Parameters></Method></Actions><ObjectPaths><Property Id="17" ParentId="15" Name="FieldLinks" /><Method Id="19" ParentId="17" Name="GetById"><Parameters><Parameter Type="Guid">{${formatting.escapeXml(args.options.id)}}</Parameter></Parameters></Method><Identity Id="15" Name="09eec89e-709b-0000-558c-c222dcaf9162|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${siteId}:web:${webId}:list:${listId}:contenttype:${formatting.escapeXml(args.options.contentTypeId)}" /></ObjectPaths></Request>`;
                }
                else {
                    requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName=".NET Library" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="77" ObjectPathId="76" /><ObjectPath Id="79" ObjectPathId="78" /><Method Name="DeleteObject" Id="80" ObjectPathId="78" /><Method Name="Update" Id="81" ObjectPathId="24"><Parameters><Parameter Type="Boolean">${updateChildContentTypes}</Parameter></Parameters></Method></Actions><ObjectPaths><Property Id="76" ParentId="24" Name="FieldLinks" /><Method Id="78" ParentId="76" Name="GetById"><Parameters><Parameter Type="Guid">{${formatting.escapeXml(args.options.id)}}</Parameter></Parameters></Method><Identity Id="24" Name="6b3ec69e-00a7-0000-55a3-61f8d779d2b3|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${siteId}:web:${webId}:contenttype:${formatting.escapeXml(args.options.contentTypeId)}" /></ObjectPaths></Request>`;
                }
                const requestOptions = {
                    url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
                    headers: {
                        'X-RequestDigest': requestDigest
                    },
                    data: requestBody
                };
                const res = await request.post(requestOptions);
                const json = JSON.parse(res);
                const response = json[0];
                if (response.ErrorInfo) {
                    throw response.ErrorInfo.ErrorMessage;
                }
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeFieldLink();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the column ${args.options.id} from content type ${args.options.contentTypeId}?` });
            if (result) {
                await removeFieldLink();
            }
        }
    }
}
_SpoContentTypeFieldRemoveCommand_instances = new WeakSet(), _SpoContentTypeFieldRemoveCommand_initTelemetry = function _SpoContentTypeFieldRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listTitle: typeof args.options.listTitle !== 'undefined',
            listId: typeof args.options.listId !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            updateChildContentTypes: !!args.options.updateChildContentTypes,
            force: !!args.options.force
        });
    });
}, _SpoContentTypeFieldRemoveCommand_initOptions = function _SpoContentTypeFieldRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-l, --listTitle [listTitle]'
    }, {
        option: '--listId [listId]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '--contentTypeId <contentTypeId>'
    }, {
        option: '-i, --id <id>'
    }, {
        option: '-c, --updateChildContentTypes'
    }, {
        option: '-f, --force'
    });
}, _SpoContentTypeFieldRemoveCommand_initValidators = function _SpoContentTypeFieldRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} is not a valid GUID`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
}, _SpoContentTypeFieldRemoveCommand_initTypes = function _SpoContentTypeFieldRemoveCommand_initTypes() {
    this.types.string.push('i', 'contentTypeId');
};
export default new SpoContentTypeFieldRemoveCommand();
//# sourceMappingURL=contenttype-field-remove.js.map