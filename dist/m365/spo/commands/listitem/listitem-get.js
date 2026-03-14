var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListItemGetCommand_instances, _SpoListItemGetCommand_initTelemetry, _SpoListItemGetCommand_initOptions, _SpoListItemGetCommand_initValidators, _SpoListItemGetCommand_initTypes, _SpoListItemGetCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListItemGetCommand extends SpoCommand {
    allowUnknownOptions() {
        return true;
    }
    get name() {
        return commands.LISTITEM_GET;
    }
    get description() {
        return 'Gets a list item from the specified list';
    }
    constructor() {
        super();
        _SpoListItemGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListItemGetCommand_instances, "m", _SpoListItemGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListItemGetCommand_instances, "m", _SpoListItemGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListItemGetCommand_instances, "m", _SpoListItemGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListItemGetCommand_instances, "m", _SpoListItemGetCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _SpoListItemGetCommand_instances, "m", _SpoListItemGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
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
        const propertiesSelect = args.options.properties ? args.options.properties.split(',') : [];
        const propertiesWithSlash = propertiesSelect.filter(item => item.includes('/'));
        const propertiesToExpand = propertiesWithSlash.map(e => e.split('/')[0]);
        const expandPropertiesArray = propertiesToExpand.filter((item, pos) => propertiesToExpand.indexOf(item) === pos);
        const fieldExpand = expandPropertiesArray.length > 0 ? `&$expand=${expandPropertiesArray.join(",")}` : ``;
        if (args.options.id) {
            requestUrl += `/items(${args.options.id})`;
        }
        else {
            requestUrl += `/GetItemByUniqueId(guid'${args.options.uniqueId}')`;
        }
        const requestOptions = {
            url: `${requestUrl}?$select=${formatting.encodeQueryParameter(propertiesSelect.join(","))}${fieldExpand}`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        try {
            const itemProperties = await request.get(requestOptions);
            if (args.options.withPermissions) {
                requestOptions.url = `${requestUrl}/RoleAssignments?$expand=Member,RoleDefinitionBindings`;
                const response = await request.get(requestOptions);
                response.value.forEach((r) => {
                    r.RoleDefinitionBindings = formatting.setFriendlyPermissions(r.RoleDefinitionBindings);
                });
                itemProperties.RoleAssignments = response.value;
            }
            delete itemProperties['ID'];
            await logger.log(itemProperties);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoListItemGetCommand_instances = new WeakSet(), _SpoListItemGetCommand_initTelemetry = function _SpoListItemGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            uniqueId: typeof args.options.uniqueId !== 'undefined',
            withPermissions: !!args.options.withPermissions
        });
    });
}, _SpoListItemGetCommand_initOptions = function _SpoListItemGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '--uniqueId [uniqueId]'
    }, {
        option: '-l, --listId [listId]'
    }, {
        option: '-t, --listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '-p, --properties [properties]'
    }, {
        option: '--withPermissions'
    });
}, _SpoListItemGetCommand_initValidators = function _SpoListItemGetCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.listId &&
            !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} in option listId is not a valid GUID`;
        }
        if (args.options.id &&
            isNaN(parseInt(args.options.id))) {
            return `${args.options.id} is not a number`;
        }
        if (args.options.uniqueId &&
            !validation.isValidGuid(args.options.uniqueId)) {
            return `${args.options.uniqueId} in option uniqueId is not a valid GUID`;
        }
        return true;
    });
}, _SpoListItemGetCommand_initTypes = function _SpoListItemGetCommand_initTypes() {
    this.types.string.push('webUrl', 'listId', 'listTitle', 'id', 'uniqueId', 'properties');
}, _SpoListItemGetCommand_initOptionSets = function _SpoListItemGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] }, { options: ['id', 'uniqueId'] });
};
export default new SpoListItemGetCommand();
//# sourceMappingURL=listitem-get.js.map