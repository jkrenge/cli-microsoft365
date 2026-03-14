var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListSensitivityLabelEnsureCommand_instances, _SpoListSensitivityLabelEnsureCommand_initTelemetry, _SpoListSensitivityLabelEnsureCommand_initOptions, _SpoListSensitivityLabelEnsureCommand_initValidators, _SpoListSensitivityLabelEnsureCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListSensitivityLabelEnsureCommand extends SpoCommand {
    get name() {
        return commands.LIST_SENSITIVITYLABEL_ENSURE;
    }
    get description() {
        return 'Applies a default sensitivity label to the specified document library';
    }
    constructor() {
        super();
        _SpoListSensitivityLabelEnsureCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListSensitivityLabelEnsureCommand_instances, "m", _SpoListSensitivityLabelEnsureCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListSensitivityLabelEnsureCommand_instances, "m", _SpoListSensitivityLabelEnsureCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListSensitivityLabelEnsureCommand_instances, "m", _SpoListSensitivityLabelEnsureCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListSensitivityLabelEnsureCommand_instances, "m", _SpoListSensitivityLabelEnsureCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const sensitivityLabelId = await this.getSensitivityLabelId(args, logger);
            if (this.verbose) {
                await logger.logToStderr(`Applying a sensitivity label ${sensitivityLabelId} to the document library...`);
            }
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
            const requestOptions = {
                url: requestUrl,
                headers: {
                    accept: 'application/json;odata=nometadata',
                    'content-type': 'application/json;odata=nometadata',
                    'if-match': '*'
                },
                data: { 'DefaultSensitivityLabelForLibrary': sensitivityLabelId },
                responseType: 'json'
            };
            await request.patch(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getSensitivityLabelId(args, logger) {
        const { id, name } = args.options;
        if (id) {
            return id;
        }
        if (this.verbose) {
            await logger.logToStderr(`Retrieving sensitivity label id of ${name}...`);
        }
        const requestOptions = {
            url: `https://graph.microsoft.com/beta/security/informationProtection/sensitivityLabels?$filter=name eq '${formatting.encodeQueryParameter(name)}'&$select=id`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const res = await request.get(requestOptions);
        const sensitivityLabelItem = res.value[0];
        if (!sensitivityLabelItem) {
            throw Error(`The specified sensitivity label does not exist`);
        }
        return sensitivityLabelItem.id;
    }
}
_SpoListSensitivityLabelEnsureCommand_instances = new WeakSet(), _SpoListSensitivityLabelEnsureCommand_initTelemetry = function _SpoListSensitivityLabelEnsureCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            name: typeof args.options.name !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined'
        });
    });
}, _SpoListSensitivityLabelEnsureCommand_initOptions = function _SpoListSensitivityLabelEnsureCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-n, --name [name]'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-t, --listTitle [listTitle]'
    }, {
        option: '-l, --listId [listId]'
    }, {
        option: '--listUrl [listUrl]'
    });
}, _SpoListSensitivityLabelEnsureCommand_initValidators = function _SpoListSensitivityLabelEnsureCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} is not a valid GUID`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
}, _SpoListSensitivityLabelEnsureCommand_initOptionSets = function _SpoListSensitivityLabelEnsureCommand_initOptionSets() {
    this.optionSets.push({ options: ['name', 'id'] }, { options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListSensitivityLabelEnsureCommand();
//# sourceMappingURL=list-sensitivitylabel-ensure.js.map