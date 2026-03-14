var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListItemRecordUndeclareCommand_instances, _SpoListItemRecordUndeclareCommand_initTelemetry, _SpoListItemRecordUndeclareCommand_initOptions, _SpoListItemRecordUndeclareCommand_initValidators, _SpoListItemRecordUndeclareCommand_initOptionSets;
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListItemRecordUndeclareCommand extends SpoCommand {
    get name() {
        return commands.LISTITEM_RECORD_UNDECLARE;
    }
    get description() {
        return 'Undeclares list item as a record';
    }
    constructor() {
        super();
        _SpoListItemRecordUndeclareCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListItemRecordUndeclareCommand_instances, "m", _SpoListItemRecordUndeclareCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListItemRecordUndeclareCommand_instances, "m", _SpoListItemRecordUndeclareCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListItemRecordUndeclareCommand_instances, "m", _SpoListItemRecordUndeclareCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListItemRecordUndeclareCommand_instances, "m", _SpoListItemRecordUndeclareCommand_initOptionSets).call(this);
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
            if (this.debug) {
                await logger.logToStderr(`getting request digest for request`);
            }
            const reqDigest = await spo.getRequestDigest(args.options.webUrl);
            const formDigestValue = reqDigest.FormDigestValue;
            const objectIdentity = await spo.getCurrentWebIdentity(args.options.webUrl, formDigestValue);
            if (this.verbose) {
                await logger.logToStderr(`Undeclare list item as a record in list ${args.options.listId || args.options.listTitle || args.options.listUrl} in site ${args.options.webUrl}...`);
            }
            const requestOptions = {
                url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'Content-Type': 'text/xml',
                    'X-RequestDigest': formDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><StaticMethod TypeId="{ea8e1356-5910-4e69-bc05-d0c30ed657fc}" Name="UndeclareItemAsRecord" Id="53"><Parameters><Parameter ObjectPathId="49" /></Parameters></StaticMethod></Actions><ObjectPaths><Identity Id="49" Name="${objectIdentity.objectIdentity}:list:${listId}:item:${args.options.listItemId},1" /></ObjectPaths></Request>`
            };
            await request.post(requestOptions);
            // REST post call doesn't return anything
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
}
_SpoListItemRecordUndeclareCommand_instances = new WeakSet(), _SpoListItemRecordUndeclareCommand_initTelemetry = function _SpoListItemRecordUndeclareCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined'
        });
    });
}, _SpoListItemRecordUndeclareCommand_initOptions = function _SpoListItemRecordUndeclareCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-i, --listItemId <listItemId>'
    }, {
        option: '-l, --listId [listId]'
    }, {
        option: '-t, --listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    });
}, _SpoListItemRecordUndeclareCommand_initValidators = function _SpoListItemRecordUndeclareCommand_initValidators() {
    this.validators.push(async (args) => {
        const id = parseInt(args.options.listItemId);
        if (isNaN(id)) {
            return `${args.options.listItemId} is not a valid list item ID`;
        }
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.listId &&
            !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} in option listId is not a valid GUID`;
        }
        return true;
    });
}, _SpoListItemRecordUndeclareCommand_initOptionSets = function _SpoListItemRecordUndeclareCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListItemRecordUndeclareCommand();
//# sourceMappingURL=listitem-record-undeclare.js.map