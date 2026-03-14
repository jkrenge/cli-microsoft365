var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListItemIsRecordCommand_instances, _SpoListItemIsRecordCommand_initTelemetry, _SpoListItemIsRecordCommand_initOptions, _SpoListItemIsRecordCommand_initValidators, _SpoListItemIsRecordCommand_initOptionSets;
import { Auth } from '../../../../Auth.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListItemIsRecordCommand extends SpoCommand {
    get name() {
        return commands.LISTITEM_ISRECORD;
    }
    get description() {
        return 'Checks if the specified list item is a record';
    }
    constructor() {
        super();
        _SpoListItemIsRecordCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListItemIsRecordCommand_instances, "m", _SpoListItemIsRecordCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListItemIsRecordCommand_instances, "m", _SpoListItemIsRecordCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListItemIsRecordCommand_instances, "m", _SpoListItemIsRecordCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListItemIsRecordCommand_instances, "m", _SpoListItemIsRecordCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const resource = Auth.getResourceFromUrl(args.options.webUrl);
        let requestUrl = `${args.options.webUrl}/_api/web`;
        if (args.options.listId) {
            requestUrl += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
        }
        else if (args.options.listTitle) {
            requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
        }
        else if (args.options.listUrl) {
            const listServerRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
            requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
        }
        let listId;
        if (this.debug) {
            await logger.logToStderr(`Retrieving access token for ${resource}...`);
        }
        try {
            if (typeof args.options.listId !== 'undefined') {
                if (this.verbose) {
                    await logger.logToStderr(`List Id passed in as an argument.`);
                }
                listId = args.options.listId;
            }
            else {
                if (this.verbose) {
                    await logger.logToStderr(`Getting list id for list ${args.options.listTitle ? args.options.listTitle : args.options.listUrl}`);
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
                await logger.logToStderr(`Getting request digest for request`);
            }
            const reqDigest = await spo.getRequestDigest(args.options.webUrl);
            const formDigestValue = reqDigest.FormDigestValue;
            const webIdentityResp = await spo.getCurrentWebIdentity(args.options.webUrl, formDigestValue);
            if (this.verbose) {
                await logger.logToStderr(`Checking if list item is a record in list ${args.options.listId ? args.options.listId : args.options.listTitle ? args.options.listTitle : args.options.listUrl} in site ${args.options.webUrl}...`);
            }
            const requestBody = this.getIsRecordRequestBody(webIdentityResp.objectIdentity, listId, args.options.id);
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
    getIsRecordRequestBody(webIdentity, listId, id) {
        const requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009">
            <Actions>
              <StaticMethod TypeId="{ea8e1356-5910-4e69-bc05-d0c30ed657fc}" Name="IsRecord" Id="1"><Parameters><Parameter ObjectPathId="14" /></Parameters></StaticMethod>
            </Actions>
            <ObjectPaths>
              <Identity Id="14" Name="${webIdentity}:list:${listId}:item:${id},1" />
            </ObjectPaths>
          </Request>`;
        return requestBody;
    }
}
_SpoListItemIsRecordCommand_instances = new WeakSet(), _SpoListItemIsRecordCommand_initTelemetry = function _SpoListItemIsRecordCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined'
        });
    });
}, _SpoListItemIsRecordCommand_initOptions = function _SpoListItemIsRecordCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-i, --id <id>'
    }, {
        option: '-l, --listId [listId]'
    }, {
        option: '-t, --listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    });
}, _SpoListItemIsRecordCommand_initValidators = function _SpoListItemIsRecordCommand_initValidators() {
    this.validators.push(async (args) => {
        const id = parseInt(args.options.id);
        if (isNaN(id)) {
            return `${args.options.id} is not a valid list item ID`;
        }
        if (id < 1) {
            return `Item ID must be a positive number`;
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
}, _SpoListItemIsRecordCommand_initOptionSets = function _SpoListItemIsRecordCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListItemIsRecordCommand();
//# sourceMappingURL=listitem-isrecord.js.map