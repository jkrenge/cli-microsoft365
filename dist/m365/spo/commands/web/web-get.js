var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoWebGetCommand_instances, _SpoWebGetCommand_initTelemetry, _SpoWebGetCommand_initOptions, _SpoWebGetCommand_initValidators;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoWebGetCommand extends SpoCommand {
    get name() {
        return commands.WEB_GET;
    }
    get description() {
        return 'Retrieve information about the specified site';
    }
    constructor() {
        super();
        _SpoWebGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoWebGetCommand_instances, "m", _SpoWebGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoWebGetCommand_instances, "m", _SpoWebGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoWebGetCommand_instances, "m", _SpoWebGetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        let url = `${args.options.url}/_api/web`;
        if (args.options.withGroups) {
            url += '?$expand=AssociatedMemberGroup,AssociatedOwnerGroup,AssociatedVisitorGroup';
        }
        const requestOptions = {
            url,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        try {
            const webProperties = await request.get(requestOptions);
            if (args.options.withPermissions) {
                requestOptions.url = `${args.options.url}/_api/web/RoleAssignments?$expand=Member,RoleDefinitionBindings`;
                const response = await request.get(requestOptions);
                response.value.forEach((r) => {
                    r.RoleDefinitionBindings = formatting.setFriendlyPermissions(r.RoleDefinitionBindings);
                });
                webProperties.RoleAssignments = response.value;
            }
            await logger.log(webProperties);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoWebGetCommand_instances = new WeakSet(), _SpoWebGetCommand_initTelemetry = function _SpoWebGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            withGroups: !!args.options.withGroups,
            withPermissions: !!args.options.withPermissions
        });
    });
}, _SpoWebGetCommand_initOptions = function _SpoWebGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --url <url>'
    }, {
        option: '--withGroups'
    }, {
        option: '--withPermissions'
    });
}, _SpoWebGetCommand_initValidators = function _SpoWebGetCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.url));
};
export default new SpoWebGetCommand();
//# sourceMappingURL=web-get.js.map