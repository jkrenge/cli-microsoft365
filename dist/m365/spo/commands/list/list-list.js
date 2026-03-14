var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListListCommand_instances, _SpoListListCommand_initOptions, _SpoListListCommand_initValidators, _SpoListListCommand_initTelemetry;
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListListCommand extends SpoCommand {
    get name() {
        return commands.LIST_LIST;
    }
    get description() {
        return 'Lists all available list in the specified site';
    }
    defaultProperties() {
        return ['Title', 'Id'];
    }
    constructor() {
        super();
        _SpoListListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListListCommand_instances, "m", _SpoListListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListListCommand_instances, "m", _SpoListListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListListCommand_instances, "m", _SpoListListCommand_initTelemetry).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving all lists in site at ${args.options.webUrl}...`);
        }
        try {
            const fieldProperties = this.formatSelectProperties(args.options.properties);
            const queryParams = [`$expand=${fieldProperties.expandProperties.join(',')}`, `$select=${fieldProperties.selectProperties.join(',')}`];
            if (args.options.filter) {
                queryParams.push(`$filter=${args.options.filter}`);
            }
            const listInstances = await odata.getAllItems(`${args.options.webUrl}/_api/web/lists?${queryParams.join('&')}`);
            await logger.log(listInstances);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    formatSelectProperties(fields) {
        const selectProperties = ['RootFolder/ServerRelativeUrl'];
        const expandProperties = ['RootFolder'];
        if (!fields) {
            selectProperties.push('*');
        }
        if (fields) {
            fields.split(',').forEach((field) => {
                const subparts = field.trim().split('/');
                if (subparts.length > 1) {
                    expandProperties.push(subparts[0]);
                }
                selectProperties.push(field.trim());
            });
        }
        return {
            selectProperties: [...new Set(selectProperties)],
            expandProperties: [...new Set(expandProperties)]
        };
    }
}
_SpoListListCommand_instances = new WeakSet(), _SpoListListCommand_initOptions = function _SpoListListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-p, --properties [properties]'
    }, {
        option: '--filter [filter]'
    });
}, _SpoListListCommand_initValidators = function _SpoListListCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.webUrl));
}, _SpoListListCommand_initTelemetry = function _SpoListListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            properties: typeof args.options.properties !== 'undefined',
            filter: typeof args.options.filter !== 'undefined'
        });
    });
};
export default new SpoListListCommand();
//# sourceMappingURL=list-list.js.map