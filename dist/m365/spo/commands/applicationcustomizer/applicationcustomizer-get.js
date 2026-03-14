var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoApplicationCustomizerGetCommand_instances, _SpoApplicationCustomizerGetCommand_initTelemetry, _SpoApplicationCustomizerGetCommand_initOptions, _SpoApplicationCustomizerGetCommand_initValidators, _SpoApplicationCustomizerGetCommand_initOptionSets, _SpoApplicationCustomizerGetCommand_initTypes;
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';
class SpoApplicationCustomizerGetCommand extends SpoCommand {
    get name() {
        return commands.APPLICATIONCUSTOMIZER_GET;
    }
    get description() {
        return 'Get an application customizer that is added to a site.';
    }
    constructor() {
        super();
        _SpoApplicationCustomizerGetCommand_instances.add(this);
        this.allowedScopes = ['All', 'Site', 'Web'];
        __classPrivateFieldGet(this, _SpoApplicationCustomizerGetCommand_instances, "m", _SpoApplicationCustomizerGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoApplicationCustomizerGetCommand_instances, "m", _SpoApplicationCustomizerGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoApplicationCustomizerGetCommand_instances, "m", _SpoApplicationCustomizerGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoApplicationCustomizerGetCommand_instances, "m", _SpoApplicationCustomizerGetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoApplicationCustomizerGetCommand_instances, "m", _SpoApplicationCustomizerGetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            const customAction = await this.getCustomAction(args.options);
            if (!args.options.clientSideComponentProperties) {
                await logger.log({
                    ...customAction,
                    Scope: this.humanizeScope(customAction.Scope)
                });
            }
            else {
                const properties = formatting.tryParseJson(customAction.ClientSideComponentProperties);
                await logger.log(properties);
            }
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
    async getCustomAction(options) {
        if (options.id) {
            const customAction = await spo.getCustomActionById(options.webUrl, options.id, options.scope);
            if (!customAction || (customAction && customAction.Location !== 'ClientSideExtension.ApplicationCustomizer')) {
                throw `No application customizer with id '${options.id}' found`;
            }
            return customAction;
        }
        const filter = options.title ? `Title eq '${formatting.encodeQueryParameter(options.title)}'` : `ClientSideComponentId eq guid'${formatting.encodeQueryParameter(options.clientSideComponentId)}'`;
        const customActions = await spo.getCustomActions(options.webUrl, options.scope, `${filter} and Location eq 'ClientSideExtension.ApplicationCustomizer'`);
        if (customActions.length === 1) {
            return customActions[0];
        }
        const identifier = options.title ? `title '${options.title}'` : `Client Side Component Id '${options.clientSideComponentId}'`;
        if (customActions.length === 0) {
            throw `No application customizer with ${identifier} found`;
        }
        else {
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', customActions);
            return await cli.handleMultipleResultsFound(`Multiple application customizers with ${identifier} found.`, resultAsKeyValuePair);
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
_SpoApplicationCustomizerGetCommand_instances = new WeakSet(), _SpoApplicationCustomizerGetCommand_initTelemetry = function _SpoApplicationCustomizerGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            title: typeof args.options.title !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            clientSideComponentId: typeof args.options.clientSideComponentId !== 'undefined',
            scope: typeof args.options.scope !== 'undefined',
            clientSideComponentProperties: !!args.options.clientSideComponentProperties
        });
    });
}, _SpoApplicationCustomizerGetCommand_initOptions = function _SpoApplicationCustomizerGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-t, --title [title]'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-c, --clientSideComponentId  [clientSideComponentId]'
    }, {
        option: '-s, --scope [scope]',
        autocomplete: this.allowedScopes
    }, {
        option: '-p, --clientSideComponentProperties'
    });
}, _SpoApplicationCustomizerGetCommand_initValidators = function _SpoApplicationCustomizerGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        if (args.options.clientSideComponentId && !validation.isValidGuid(args.options.clientSideComponentId)) {
            return `${args.options.clientSideComponentId} is not a valid GUID`;
        }
        if (args.options.scope && this.allowedScopes.indexOf(args.options.scope) === -1) {
            return `'${args.options.scope}' is not a valid application customizer scope. Allowed values are: ${this.allowedScopes.join(',')}`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
}, _SpoApplicationCustomizerGetCommand_initOptionSets = function _SpoApplicationCustomizerGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['title', 'id', 'clientSideComponentId'] });
}, _SpoApplicationCustomizerGetCommand_initTypes = function _SpoApplicationCustomizerGetCommand_initTypes() {
    this.types.string.push('webUrl', 'title', 'id', 'clientSideComponentId', 'scope');
    this.types.boolean.push('clientSideComponentProperties');
};
export default new SpoApplicationCustomizerGetCommand();
//# sourceMappingURL=applicationcustomizer-get.js.map