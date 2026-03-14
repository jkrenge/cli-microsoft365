var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoCustomActionGetCommand_instances, _SpoCustomActionGetCommand_initTelemetry, _SpoCustomActionGetCommand_initOptions, _SpoCustomActionGetCommand_initValidators, _SpoCustomActionGetCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoCustomActionGetCommand extends SpoCommand {
    get name() {
        return commands.CUSTOMACTION_GET;
    }
    get description() {
        return 'Gets details for the specified custom action';
    }
    constructor() {
        super();
        _SpoCustomActionGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoCustomActionGetCommand_instances, "m", _SpoCustomActionGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoCustomActionGetCommand_instances, "m", _SpoCustomActionGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoCustomActionGetCommand_instances, "m", _SpoCustomActionGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoCustomActionGetCommand_instances, "m", _SpoCustomActionGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const customAction = await this.getCustomAction(args.options);
            if (customAction) {
                await logger.log({
                    ClientSideComponentId: customAction.ClientSideComponentId,
                    ClientSideComponentProperties: customAction.ClientSideComponentProperties,
                    CommandUIExtension: customAction.CommandUIExtension,
                    Description: customAction.Description,
                    Group: customAction.Group,
                    Id: customAction.Id,
                    ImageUrl: customAction.ImageUrl,
                    Location: customAction.Location,
                    Name: customAction.Name,
                    RegistrationId: customAction.RegistrationId,
                    RegistrationType: customAction.RegistrationType,
                    Rights: JSON.stringify(customAction.Rights),
                    Scope: this.humanizeScope(customAction.Scope),
                    ScriptBlock: customAction.ScriptBlock,
                    ScriptSrc: customAction.ScriptSrc,
                    Sequence: customAction.Sequence,
                    Title: customAction.Title,
                    Url: customAction.Url,
                    VersionOfUserCustomAction: customAction.VersionOfUserCustomAction
                });
            }
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
    async getCustomAction(options) {
        if (options.id) {
            const customAction = await spo.getCustomActionById(options.webUrl, options.id, options.scope);
            if (!customAction) {
                throw `No user custom action with id '${options.id}' found`;
            }
            return customAction;
        }
        else if (options.title) {
            const customActions = await spo.getCustomActions(options.webUrl, options.scope, `Title eq '${formatting.encodeQueryParameter(options.title)}'`);
            if (customActions.length === 1) {
                return customActions[0];
            }
            if (customActions.length === 0) {
                throw `No user custom action with title '${options.title}' found`;
            }
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', customActions);
            return await cli.handleMultipleResultsFound(`Multiple user custom actions with title '${options.title}' found.`, resultAsKeyValuePair);
        }
        else {
            const customActions = await spo.getCustomActions(options.webUrl, options.scope, `ClientSideComponentId eq guid'${options.clientSideComponentId}'`);
            if (customActions.length === 0) {
                throw `No user custom action with ClientSideComponentId '${options.clientSideComponentId}' found`;
            }
            if (customActions.length > 1) {
                const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', customActions);
                return await cli.handleMultipleResultsFound(`Multiple user custom actions with ClientSideComponentId '${options.clientSideComponentId}' found.`, resultAsKeyValuePair);
            }
            return customActions[0];
        }
    }
    humanizeScope(scope) {
        switch (scope) {
            case 2:
                return "Site";
            case 3:
                return "Web";
        }
        return `${scope}`;
    }
}
_SpoCustomActionGetCommand_instances = new WeakSet(), _SpoCustomActionGetCommand_initTelemetry = function _SpoCustomActionGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            title: typeof args.options.title !== 'undefined',
            clientSideComponentId: typeof args.options.clientSideComponentId !== 'undefined',
            scope: args.options.scope || 'All'
        });
    });
}, _SpoCustomActionGetCommand_initOptions = function _SpoCustomActionGetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-t, --title [title]'
    }, {
        option: '-c, --clientSideComponentId [clientSideComponentId]'
    }, {
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-s, --scope [scope]',
        autocomplete: ['Site', 'Web', 'All']
    });
}, _SpoCustomActionGetCommand_initValidators = function _SpoCustomActionGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && validation.isValidGuid(args.options.id) === false) {
            return `${args.options.id} is not valid. Custom action id (Guid) expected.`;
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
        if (args.options.clientSideComponentId && validation.isValidGuid(args.options.clientSideComponentId) === false) {
            return `${args.options.clientSideComponentId} is not a valid GUID.`;
        }
        return true;
    });
}, _SpoCustomActionGetCommand_initOptionSets = function _SpoCustomActionGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'title', 'clientSideComponentId'] });
};
export default new SpoCustomActionGetCommand();
//# sourceMappingURL=customaction-get.js.map