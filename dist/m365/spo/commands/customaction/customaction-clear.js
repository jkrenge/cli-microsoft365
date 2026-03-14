var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoCustomActionClearCommand_instances, _SpoCustomActionClearCommand_initTelemetry, _SpoCustomActionClearCommand_initOptions, _SpoCustomActionClearCommand_initValidators;
import chalk from 'chalk';
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoCustomActionClearCommand extends SpoCommand {
    get name() {
        return commands.CUSTOMACTION_CLEAR;
    }
    get description() {
        return 'Deletes all custom actions in the collection';
    }
    constructor() {
        super();
        _SpoCustomActionClearCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoCustomActionClearCommand_instances, "m", _SpoCustomActionClearCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoCustomActionClearCommand_instances, "m", _SpoCustomActionClearCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoCustomActionClearCommand_instances, "m", _SpoCustomActionClearCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const clearCustomActions = async () => {
            try {
                if (args.options.scope && args.options.scope.toLowerCase() !== "all") {
                    await this.clearScopedCustomActions(args.options);
                }
                else {
                    await this.clearAllScopes(args.options);
                }
            }
            catch (err) {
                this.handleRejectedPromise(err);
            }
        };
        if (args.options.force) {
            await clearCustomActions();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to clear all the user custom actions with scope ${chalk.yellow(args.options.scope || 'All')}?` });
            if (result) {
                await clearCustomActions();
            }
        }
    }
    clearScopedCustomActions(options) {
        const requestOptions = {
            url: `${options.webUrl}/_api/${options.scope}/UserCustomActions/clear`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        return request.post(requestOptions);
    }
    /**
     * Clear request with `web` scope is send first.
     * Another clear request is send with `site` scope after.
     */
    async clearAllScopes(options) {
        options.scope = "Web";
        await this.clearScopedCustomActions(options);
        options.scope = "Site";
        await this.clearScopedCustomActions(options);
    }
}
_SpoCustomActionClearCommand_instances = new WeakSet(), _SpoCustomActionClearCommand_initTelemetry = function _SpoCustomActionClearCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            scope: args.options.scope || 'All',
            force: (!(!args.options.force)).toString()
        });
    });
}, _SpoCustomActionClearCommand_initOptions = function _SpoCustomActionClearCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-s, --scope [scope]',
        autocomplete: ['Site', 'Web', 'All']
    }, {
        option: '-f, --force'
    });
}, _SpoCustomActionClearCommand_initValidators = function _SpoCustomActionClearCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (typeof isValidUrl === 'string') {
            return isValidUrl;
        }
        if (args.options.scope &&
            args.options.scope !== 'Site' &&
            args.options.scope !== 'Web' &&
            args.options.scope !== 'All') {
            return `${args.options.scope} is not a valid custom action scope. Allowed values are Site|Web|All`;
        }
        return true;
    });
};
export default new SpoCustomActionClearCommand();
//# sourceMappingURL=customaction-clear.js.map