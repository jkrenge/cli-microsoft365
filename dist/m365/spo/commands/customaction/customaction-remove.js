var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoCustomActionRemoveCommand_instances, _SpoCustomActionRemoveCommand_initTelemetry, _SpoCustomActionRemoveCommand_initOptions, _SpoCustomActionRemoveCommand_initValidators, _SpoCustomActionRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoCustomActionRemoveCommand extends SpoCommand {
    get name() {
        return commands.CUSTOMACTION_REMOVE;
    }
    get description() {
        return 'Removes the specified custom action';
    }
    constructor() {
        super();
        _SpoCustomActionRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoCustomActionRemoveCommand_instances, "m", _SpoCustomActionRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoCustomActionRemoveCommand_instances, "m", _SpoCustomActionRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoCustomActionRemoveCommand_instances, "m", _SpoCustomActionRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoCustomActionRemoveCommand_instances, "m", _SpoCustomActionRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const removeCustomAction = async () => {
            try {
                let customAction;
                if (args.options.scope && args.options.scope.toLowerCase() !== "all") {
                    customAction = await this.removeScopedCustomAction(args.options);
                }
                else {
                    customAction = await this.searchAllScopes(args.options);
                }
                if (this.verbose) {
                    if (customAction && customAction["odata.null"] === true) {
                        await logger.logToStderr(`Custom action with id ${args.options.id} not found`);
                    }
                }
            }
            catch (err) {
                this.handleRejectedPromise(err);
            }
        };
        if (args.options.force) {
            await removeCustomAction();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the ${args.options.id} user custom action?` });
            if (result) {
                await removeCustomAction();
            }
        }
    }
    async getCustomActionId(options) {
        if (options.id) {
            return options.id;
        }
        const customActions = await spo.getCustomActions(options.webUrl, options.scope, `Title eq '${formatting.encodeQueryParameter(options.title)}'`);
        if (customActions.length === 1) {
            return customActions[0].Id;
        }
        if (customActions.length === 0) {
            throw `No user custom action with title '${options.title}' found`;
        }
        const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', customActions);
        const result = await cli.handleMultipleResultsFound(`Multiple user custom actions with title '${options.title}' found.`, resultAsKeyValuePair);
        return result.Id;
    }
    async removeScopedCustomAction(options) {
        const customActionId = await this.getCustomActionId(options);
        const requestOptions = {
            url: `${options.webUrl}/_api/${options.scope}/UserCustomActions('${formatting.encodeQueryParameter(customActionId)}')')`,
            headers: {
                accept: 'application/json;odata=nometadata',
                'X-HTTP-Method': 'DELETE'
            },
            responseType: 'json'
        };
        return await request.post(requestOptions);
    }
    /**
     * Remove request with `web` scope is send first.
     * If custom action not found then
     * another get request is send with `site` scope.
     */
    async searchAllScopes(options) {
        options.scope = "Web";
        const webResult = await this.removeScopedCustomAction(options);
        if (!webResult) {
            return webResult;
        }
        options.scope = "Site";
        const siteResult = await this.removeScopedCustomAction(options);
        return siteResult;
    }
}
_SpoCustomActionRemoveCommand_instances = new WeakSet(), _SpoCustomActionRemoveCommand_initTelemetry = function _SpoCustomActionRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            scope: args.options.scope || 'All',
            force: (!(!args.options.force)).toString()
        });
    });
}, _SpoCustomActionRemoveCommand_initOptions = function _SpoCustomActionRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-t, --title [title]'
    }, {
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-s, --scope [scope]',
        autocomplete: ['Site', 'Web', 'All']
    }, {
        option: '-f, --force'
    });
}, _SpoCustomActionRemoveCommand_initValidators = function _SpoCustomActionRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && validation.isValidGuid(args.options.id) === false) {
            return `${args.options.id} is not valid. Custom action Id (GUID) expected.`;
        }
        const isValidUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (typeof isValidUrl === 'string') {
            return isValidUrl;
        }
        if (args.options.scope) {
            if (args.options.scope !== 'Site' &&
                args.options.scope !== 'Web' &&
                args.options.scope !== 'All') {
                return `${args.options.scope} is not a valid custom action scope. Allowed values are Site|Web|All`;
            }
        }
        return true;
    });
}, _SpoCustomActionRemoveCommand_initOptionSets = function _SpoCustomActionRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'title'] });
};
export default new SpoCustomActionRemoveCommand();
//# sourceMappingURL=customaction-remove.js.map