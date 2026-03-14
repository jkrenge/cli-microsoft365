var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraM365GroupGetCommand_instances, _EntraM365GroupGetCommand_initTelemetry, _EntraM365GroupGetCommand_initOptions, _EntraM365GroupGetCommand_initValidators, _EntraM365GroupGetCommand_initOptionSets, _EntraM365GroupGetCommand_initTypes;
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraM365GroupGetCommand extends GraphCommand {
    get name() {
        return commands.M365GROUP_GET;
    }
    get description() {
        return 'Gets information about the specified Microsoft 365 Group or Microsoft Teams team';
    }
    constructor() {
        super();
        _EntraM365GroupGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraM365GroupGetCommand_instances, "m", _EntraM365GroupGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupGetCommand_instances, "m", _EntraM365GroupGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupGetCommand_instances, "m", _EntraM365GroupGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupGetCommand_instances, "m", _EntraM365GroupGetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupGetCommand_instances, "m", _EntraM365GroupGetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        let group;
        try {
            if (args.options.id) {
                group = await entraGroup.getGroupById(args.options.id);
            }
            else {
                group = await entraGroup.getGroupByDisplayName(args.options.displayName);
            }
            const isUnifiedGroup = await entraGroup.isUnifiedGroup(group.id);
            if (!isUnifiedGroup) {
                throw Error(`Specified group with id '${group.id}' is not a Microsoft 365 group.`);
            }
            const requestExtendedOptions = {
                url: `${this.resource}/v1.0/groups/${group.id}?$select=allowExternalSenders,autoSubscribeNewMembers,hideFromAddressLists,hideFromOutlookClients,isSubscribedByMail`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            const groupExtended = await request.get(requestExtendedOptions);
            group = { ...group, ...groupExtended };
            if (args.options.withSiteUrl) {
                const requestOptions = {
                    url: `${this.resource}/v1.0/groups/${group.id}/drive?$select=webUrl`,
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    },
                    responseType: 'json'
                };
                const res = await request.get(requestOptions);
                group.siteUrl = res.webUrl ? res.webUrl.substring(0, res.webUrl.lastIndexOf('/')) : '';
            }
            await logger.log(group);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_EntraM365GroupGetCommand_instances = new WeakSet(), _EntraM365GroupGetCommand_initTelemetry = function _EntraM365GroupGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            withSiteUrl: !!args.options.withSiteUrl
        });
    });
}, _EntraM365GroupGetCommand_initOptions = function _EntraM365GroupGetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-n, --displayName [displayName]'
    }, {
        option: '--withSiteUrl'
    });
}, _EntraM365GroupGetCommand_initValidators = function _EntraM365GroupGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
}, _EntraM365GroupGetCommand_initOptionSets = function _EntraM365GroupGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'displayName'] });
}, _EntraM365GroupGetCommand_initTypes = function _EntraM365GroupGetCommand_initTypes() {
    this.types.string.push('id', 'displayName');
};
export default new EntraM365GroupGetCommand();
//# sourceMappingURL=m365group-get.js.map