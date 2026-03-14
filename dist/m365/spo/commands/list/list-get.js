var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListGetCommand_instances, _SpoListGetCommand_initTelemetry, _SpoListGetCommand_initOptions, _SpoListGetCommand_initValidators, _SpoListGetCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { DefaultTrimModeType } from "./ListInstance.js";
import { ListPrincipalType } from './ListPrincipalType.js';
class SpoListGetCommand extends SpoCommand {
    get name() {
        return commands.LIST_GET;
    }
    get description() {
        return 'Gets information about the specific list';
    }
    constructor() {
        super();
        _SpoListGetCommand_instances.add(this);
        this.supportedBaseTemplates = [101, 109, 110, 111, 113, 114, 115, 116, 117, 119, 121, 122, 123, 126, 130, 175];
        __classPrivateFieldGet(this, _SpoListGetCommand_instances, "m", _SpoListGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListGetCommand_instances, "m", _SpoListGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListGetCommand_instances, "m", _SpoListGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListGetCommand_instances, "m", _SpoListGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving information for list in site at ${args.options.webUrl}...`);
        }
        let requestUrl = `${args.options.webUrl}/_api/web/`;
        if (args.options.id) {
            requestUrl += `lists(guid'${formatting.encodeQueryParameter(args.options.id)}')`;
        }
        else if (args.options.title) {
            requestUrl += `lists/GetByTitle('${formatting.encodeQueryParameter(args.options.title)}')`;
        }
        else if (args.options.url) {
            const listServerRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.url);
            requestUrl += `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
        }
        else if (args.options.default) {
            requestUrl += `DefaultDocumentLibrary`;
        }
        const fieldsProperties = this.formatSelectProperties(args.options.properties, args.options.withPermissions);
        const queryParams = [];
        if (fieldsProperties.selectProperties.length > 0) {
            queryParams.push(`$select=${fieldsProperties.selectProperties.join(',')}`);
        }
        if (fieldsProperties.expandProperties.length > 0) {
            queryParams.push(`$expand=${fieldsProperties.expandProperties.join(',')}`);
        }
        const requestOptions = {
            url: `${requestUrl}${queryParams.length > 0 ? `?${queryParams.join('&')}` : ''}`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        try {
            const listInstance = await request.get(requestOptions);
            if (args.options.withPermissions) {
                listInstance.RoleAssignments.forEach(r => {
                    r.Member.PrincipalTypeString = ListPrincipalType[r.Member.PrincipalType];
                });
            }
            if (this.supportedBaseTemplates.some(template => template === listInstance.BaseTemplate)) {
                await this.retrieveVersionPolicies(requestUrl, listInstance);
            }
            if (listInstance.VersionPolicies) {
                listInstance.VersionPolicies.DefaultTrimModeValue = DefaultTrimModeType[listInstance.VersionPolicies.DefaultTrimMode];
            }
            await logger.log(listInstance);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    formatSelectProperties(properties, withPermissions) {
        const selectProperties = [];
        let expandProperties = [];
        if (withPermissions) {
            expandProperties = ['HasUniqueRoleAssignments', 'RoleAssignments/Member', 'RoleAssignments/RoleDefinitionBindings', 'VersionPolicies'];
        }
        if (properties) {
            properties.split(',').forEach((property) => {
                const subparts = property.trim().split('/');
                if (subparts.length > 1) {
                    expandProperties.push(subparts[0]);
                }
                selectProperties.push(property.trim());
            });
        }
        return {
            selectProperties: [...new Set(selectProperties)],
            expandProperties: [...new Set(expandProperties)]
        };
    }
    async retrieveVersionPolicies(requestUrl, listInstance) {
        const requestOptions = {
            url: `${requestUrl}?$select=VersionPolicies&$expand=VersionPolicies`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const { VersionPolicies } = await request.get(requestOptions);
        listInstance.VersionPolicies = VersionPolicies;
        listInstance.VersionPolicies.DefaultTrimModeValue = DefaultTrimModeType[listInstance.VersionPolicies.DefaultTrimMode];
    }
}
_SpoListGetCommand_instances = new WeakSet(), _SpoListGetCommand_initTelemetry = function _SpoListGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            title: typeof args.options.title !== 'undefined',
            url: typeof args.options.url !== 'undefined',
            properties: typeof args.options.properties !== 'undefined',
            withPermissions: !!args.options.withPermissions,
            default: !!args.options.default
        });
    });
}, _SpoListGetCommand_initOptions = function _SpoListGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-t, --title [title]'
    }, {
        option: '--url [url]'
    }, {
        option: '--default'
    }, {
        option: '-p, --properties [properties]'
    }, {
        option: '--withPermissions'
    });
}, _SpoListGetCommand_initValidators = function _SpoListGetCommand_initValidators() {
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
        return true;
    });
}, _SpoListGetCommand_initOptionSets = function _SpoListGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'title', 'url', 'default'] });
};
export default new SpoListGetCommand();
//# sourceMappingURL=list-get.js.map