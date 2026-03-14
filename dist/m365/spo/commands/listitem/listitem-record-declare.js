var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListItemRecordDeclareCommand_instances, _SpoListItemRecordDeclareCommand_initTelemetry, _SpoListItemRecordDeclareCommand_initOptions, _SpoListItemRecordDeclareCommand_initValidators, _SpoListItemRecordDeclareCommand_initOptionSets;
import config from "../../../../config.js";
import request from '../../../../request.js';
import { formatting } from "../../../../utils/formatting.js";
import { spo } from "../../../../utils/spo.js";
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from "../../../../utils/validation.js";
import SpoCommand from "../../../base/SpoCommand.js";
import commands from "../../commands.js";
class SpoListItemRecordDeclareCommand extends SpoCommand {
    get name() {
        return commands.LISTITEM_RECORD_DECLARE;
    }
    get description() {
        return "Declares the specified list item as a record";
    }
    constructor() {
        super();
        _SpoListItemRecordDeclareCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListItemRecordDeclareCommand_instances, "m", _SpoListItemRecordDeclareCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListItemRecordDeclareCommand_instances, "m", _SpoListItemRecordDeclareCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListItemRecordDeclareCommand_instances, "m", _SpoListItemRecordDeclareCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListItemRecordDeclareCommand_instances, "m", _SpoListItemRecordDeclareCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            let listId = '';
            if (args.options.listId) {
                listId = args.options.listId;
            }
            else {
                let requestUrl = `${args.options.webUrl}/_api/web`;
                if (args.options.listTitle) {
                    requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
                }
                else if (args.options.listUrl) {
                    const listServerRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
                    requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
                }
                const requestOptions = {
                    url: `${requestUrl}?$select=Id`,
                    headers: {
                        accept: 'application/json;odata=nometadata'
                    },
                    responseType: 'json'
                };
                const list = await request.get(requestOptions);
                listId = list.Id;
            }
            const contextResponse = await spo.getRequestDigest(args.options.webUrl);
            const formDigestValue = contextResponse.FormDigestValue;
            const webIdentityResp = await spo.getCurrentWebIdentity(args.options.webUrl, formDigestValue);
            const webIdentity = webIdentityResp.objectIdentity;
            const requestBody = this.getDeclareRecordRequestBody(webIdentity, listId, args.options.listItemId, args.options.date || '');
            const requestOptions = {
                url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'Content-Type': 'text/xml',
                    'X-RequestDigest': formDigestValue
                },
                data: requestBody
            };
            const res = await request.post(requestOptions);
            const json = JSON.parse(res);
            const response = json[0];
            if (response.ErrorInfo) {
                throw response.ErrorInfo.ErrorMessage;
            }
            else {
                const result = json[json.length - 1];
                await logger.log(result);
            }
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
    getDeclareRecordRequestBody(webIdentity, listId, id, date) {
        let requestBody;
        if (date.length === 10) {
            requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><StaticMethod TypeId="{ea8e1356-5910-4e69-bc05-d0c30ed657fc}" Name="DeclareItemAsRecordWithDeclarationDate" Id="48"><Parameters><Parameter ObjectPathId="21" /><Parameter Type="DateTime">${date}</Parameter></Parameters></StaticMethod></Actions><ObjectPaths><Identity Id="21" Name="${webIdentity}:list:${listId}:item:${id},1" /></ObjectPaths></Request>`;
        }
        else {
            requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><StaticMethod TypeId="{ea8e1356-5910-4e69-bc05-d0c30ed657fc}" Name="DeclareItemAsRecord" Id="37"><Parameters><Parameter ObjectPathId="12" /></Parameters></StaticMethod></Actions><ObjectPaths><Identity Id="12" Name="${webIdentity}:list:${listId}:item:${id},1" /></ObjectPaths></Request>`;
        }
        return requestBody;
    }
}
_SpoListItemRecordDeclareCommand_instances = new WeakSet(), _SpoListItemRecordDeclareCommand_initTelemetry = function _SpoListItemRecordDeclareCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            date: typeof args.options.date !== 'undefined'
        });
    });
}, _SpoListItemRecordDeclareCommand_initOptions = function _SpoListItemRecordDeclareCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-l, --listId [listId]'
    }, {
        option: '-t, --listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '-i, --listItemId <listItemId>'
    }, {
        option: '-d, --date [date]'
    });
}, _SpoListItemRecordDeclareCommand_initValidators = function _SpoListItemRecordDeclareCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} in option listId is not a valid GUID`;
        }
        const id = parseInt(args.options.listItemId);
        if (isNaN(id)) {
            return `${args.options.listItemId} is not a number`;
        }
        if (id < 1) {
            return `Item ID must be a positive number`;
        }
        if (args.options.date && !validation.isValidISODate(args.options.date)) {
            return `${args.options.date} in option date is not in ISO format (yyyy-mm-dd)`;
        }
        return true;
    });
}, _SpoListItemRecordDeclareCommand_initOptionSets = function _SpoListItemRecordDeclareCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListItemRecordDeclareCommand();
//# sourceMappingURL=listitem-record-declare.js.map