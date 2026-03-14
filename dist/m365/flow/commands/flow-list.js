var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _FlowListCommand_instances, _FlowListCommand_initTelemetry, _FlowListCommand_initOptions, _FlowListCommand_initValidators, _FlowListCommand_initTypes;
import { formatting } from '../../../utils/formatting.js';
import { odata } from '../../../utils/odata.js';
import PowerAutomateCommand from '../../base/PowerAutomateCommand.js';
import commands from '../commands.js';
class FlowListCommand extends PowerAutomateCommand {
    get name() {
        return commands.LIST;
    }
    get description() {
        return 'Lists Power Automate flows in the given environment';
    }
    defaultProperties() {
        return ['name', 'displayName'];
    }
    constructor() {
        super();
        _FlowListCommand_instances.add(this);
        this.allowedSharingStatuses = ['all', 'personal', 'ownedByMe', 'sharedWithMe'];
        __classPrivateFieldGet(this, _FlowListCommand_instances, "m", _FlowListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _FlowListCommand_instances, "m", _FlowListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _FlowListCommand_instances, "m", _FlowListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _FlowListCommand_instances, "m", _FlowListCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Getting Power Automate flows${args.options.asAdmin ? ' as admin' : ''} in environment '${args.options.environmentName}'...`);
        }
        try {
            const { environmentName, asAdmin, sharingStatus, withSolutions } = args.options;
            let items = [];
            if (sharingStatus === 'personal') {
                const url = this.getApiUrl(environmentName, asAdmin, withSolutions, 'personal');
                items = await odata.getAllItems(url);
            }
            else if (sharingStatus === 'sharedWithMe') {
                const url = this.getApiUrl(environmentName, asAdmin, withSolutions, 'team');
                items = await odata.getAllItems(url);
            }
            else if (sharingStatus === 'all') {
                let url = this.getApiUrl(environmentName, asAdmin, withSolutions, 'personal');
                items = await odata.getAllItems(url);
                url = this.getApiUrl(environmentName, asAdmin, withSolutions, 'team');
                const teamFlows = await odata.getAllItems(url);
                items = items.concat(teamFlows);
            }
            else {
                const url = this.getApiUrl(environmentName, asAdmin, withSolutions);
                items = await odata.getAllItems(url);
            }
            // Remove duplicates
            items = items.filter((flow, index, self) => index === self.findIndex(f => f.id === flow.id));
            if (args.options.output && args.options.output !== 'json') {
                items.forEach(flow => {
                    flow.displayName = flow.properties.displayName;
                });
            }
            await logger.log(items);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getApiUrl(environmentName, asAdmin, includeSolutionFlows, filter) {
        const baseEndpoint = `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple`;
        const environmentSegment = `/environments/${formatting.encodeQueryParameter(environmentName)}`;
        const adminSegment = `/scopes/admin${environmentSegment}/v2`;
        const flowsEndpoint = '/flows?api-version=2016-11-01';
        const filterQuery = filter === 'personal' || filter === 'team' ? `&$filter=search('${filter}')` : '';
        const includeQuery = includeSolutionFlows ? '&include=includeSolutionCloudFlows' : '';
        return `${baseEndpoint}${asAdmin ? adminSegment : environmentSegment}${flowsEndpoint}${filterQuery}${includeQuery}`;
    }
}
_FlowListCommand_instances = new WeakSet(), _FlowListCommand_initTelemetry = function _FlowListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            sharingStatus: typeof args.options.sharingStatus !== 'undefined',
            withSolutions: !!args.options.withSolutions,
            asAdmin: !!args.options.asAdmin
        });
    });
}, _FlowListCommand_initOptions = function _FlowListCommand_initOptions() {
    this.options.unshift({
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '--sharingStatus [sharingStatus]',
        autocomplete: this.allowedSharingStatuses
    }, {
        option: '--withSolutions'
    }, {
        option: '--asAdmin'
    });
}, _FlowListCommand_initValidators = function _FlowListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.asAdmin && args.options.sharingStatus) {
            return `The options asAdmin and sharingStatus cannot be specified together.`;
        }
        if (args.options.sharingStatus && !this.allowedSharingStatuses.some(status => status === args.options.sharingStatus)) {
            return `${args.options.sharingStatus} is not a valid sharing status. Allowed values are: ${this.allowedSharingStatuses.join(', ')}`;
        }
        return true;
    });
}, _FlowListCommand_initTypes = function _FlowListCommand_initTypes() {
    this.types.string.push('environmentName', 'sharingStatus');
    this.types.boolean.push('withSolutions', 'asAdmin');
};
export default new FlowListCommand();
//# sourceMappingURL=flow-list.js.map