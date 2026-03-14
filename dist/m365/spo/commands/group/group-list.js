var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoGroupListCommand_instances, _SpoGroupListCommand_initTelemetry, _SpoGroupListCommand_initOptions, _SpoGroupListCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoGroupListCommand extends SpoCommand {
    get name() {
        return commands.GROUP_LIST;
    }
    get description() {
        return 'Lists all the groups within specific web';
    }
    defaultProperties() {
        return ['Id', 'Title', 'LoginName', 'IsHiddenInUI', 'PrincipalType', 'Type'];
    }
    constructor() {
        super();
        _SpoGroupListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoGroupListCommand_instances, "m", _SpoGroupListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoGroupListCommand_instances, "m", _SpoGroupListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoGroupListCommand_instances, "m", _SpoGroupListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving list of groups for specified web at ${args.options.webUrl}...`);
        }
        const baseUrl = `${args.options.webUrl}/_api/web`;
        try {
            if (!args.options.associatedGroupsOnly) {
                await this.getSiteGroups(baseUrl, logger);
            }
            else {
                await this.getAssociatedGroups(baseUrl, args.options, logger);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getSiteGroups(baseUrl, logger) {
        const groupProperties = await odata.getAllItems(`${baseUrl}/sitegroups`);
        await logger.log(groupProperties);
    }
    async getAssociatedGroups(baseUrl, options, logger) {
        const requestOptions = {
            url: baseUrl + '?$expand=AssociatedOwnerGroup,AssociatedMemberGroup,AssociatedVisitorGroup&$select=AssociatedOwnerGroup,AssociatedMemberGroup,AssociatedVisitorGroup',
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const groupProperties = await request.get(requestOptions);
        if (!options.output || !cli.shouldTrimOutput(options.output)) {
            await logger.log(groupProperties);
        }
        else {
            //converted to text friendly output
            const output = Object.getOwnPropertyNames(groupProperties).map(prop => ({ Type: prop, ...groupProperties[prop] }));
            await logger.log(output);
        }
    }
}
_SpoGroupListCommand_instances = new WeakSet(), _SpoGroupListCommand_initTelemetry = function _SpoGroupListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            associatedGroupsOnly: (!(!args.options.associatedGroupsOnly)).toString()
        });
    });
}, _SpoGroupListCommand_initOptions = function _SpoGroupListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--associatedGroupsOnly'
    });
}, _SpoGroupListCommand_initValidators = function _SpoGroupListCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.webUrl));
};
export default new SpoGroupListCommand();
//# sourceMappingURL=group-list.js.map