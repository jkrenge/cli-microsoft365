var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoApplicationCustomizerRemoveCommand_instances, _SpoApplicationCustomizerRemoveCommand_initOptions, _SpoApplicationCustomizerRemoveCommand_initTelemetry, _SpoApplicationCustomizerRemoveCommand_initValidators, _SpoApplicationCustomizerRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { spo } from '../../../../utils/spo.js';
import { formatting } from '../../../../utils/formatting.js';
class SpoApplicationCustomizerRemoveCommand extends SpoCommand {
    get name() {
        return commands.APPLICATIONCUSTOMIZER_REMOVE;
    }
    get description() {
        return 'Removes an application customizer that is added to a site';
    }
    constructor() {
        super();
        _SpoApplicationCustomizerRemoveCommand_instances.add(this);
        this.allowedScopes = ['Site', 'Web', 'All'];
        __classPrivateFieldGet(this, _SpoApplicationCustomizerRemoveCommand_instances, "m", _SpoApplicationCustomizerRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoApplicationCustomizerRemoveCommand_instances, "m", _SpoApplicationCustomizerRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoApplicationCustomizerRemoveCommand_instances, "m", _SpoApplicationCustomizerRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoApplicationCustomizerRemoveCommand_instances, "m", _SpoApplicationCustomizerRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (args.options.force) {
                return await this.removeApplicationCustomizer(logger, args.options);
            }
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the application customizer '${args.options.clientSideComponentId || args.options.title || args.options.id}'?` });
            if (result) {
                await this.removeApplicationCustomizer(logger, args.options);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async removeApplicationCustomizer(logger, options) {
        const applicationCustomizer = await this.getApplicationCustomizer(options);
        if (this.verbose) {
            await logger.logToStderr(`Removing application customizer '${options.clientSideComponentId || options.title || options.id}' from the site '${options.webUrl}'...`);
        }
        const requestOptions = {
            url: `${options.webUrl}/_api/${applicationCustomizer.Scope.toString() === '2' ? 'Site' : 'Web'}/UserCustomActions('${applicationCustomizer.Id}')`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        await request.delete(requestOptions);
    }
    async getApplicationCustomizer(options) {
        const resolvedScope = options.scope || 'All';
        let appCustomizers = [];
        if (options.id) {
            const appCustomizer = await spo.getCustomActionById(options.webUrl, options.id, resolvedScope);
            if (appCustomizer) {
                appCustomizers.push(appCustomizer);
            }
        }
        else if (options.title) {
            appCustomizers = await spo.getCustomActions(options.webUrl, resolvedScope, `(Title eq '${formatting.encodeQueryParameter(options.title)}') and (startswith(Location,'ClientSideExtension.ApplicationCustomizer'))`);
        }
        else {
            appCustomizers = await spo.getCustomActions(options.webUrl, resolvedScope, `(ClientSideComponentId eq guid'${options.clientSideComponentId}') and (startswith(Location,'ClientSideExtension.ApplicationCustomizer'))`);
        }
        if (appCustomizers.length === 0) {
            throw `No application customizer with ${options.title && `title '${options.title}'` || options.clientSideComponentId && `ClientSideComponentId '${options.clientSideComponentId}'` || options.id && `id '${options.id}'`} found`;
        }
        if (appCustomizers.length > 1) {
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', appCustomizers);
            return await cli.handleMultipleResultsFound(`Multiple application customizer with ${options.title ? `title '${options.title}'` : `ClientSideComponentId '${options.clientSideComponentId}'`} found.`, resultAsKeyValuePair);
        }
        return appCustomizers[0];
    }
}
_SpoApplicationCustomizerRemoveCommand_instances = new WeakSet(), _SpoApplicationCustomizerRemoveCommand_initOptions = function _SpoApplicationCustomizerRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-t, --title [title]'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-c, --clientSideComponentId [clientSideComponentId]'
    }, {
        option: '-s, --scope [scope]', autocomplete: this.allowedScopes
    }, {
        option: '-f, --force'
    });
}, _SpoApplicationCustomizerRemoveCommand_initTelemetry = function _SpoApplicationCustomizerRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            title: typeof args.options.title !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            clientSideComponentId: typeof args.options.clientSideComponentId !== 'undefined',
            scope: typeof args.options.scope !== 'undefined',
            force: !!args.options.force
        });
    });
}, _SpoApplicationCustomizerRemoveCommand_initValidators = function _SpoApplicationCustomizerRemoveCommand_initValidators() {
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
}, _SpoApplicationCustomizerRemoveCommand_initOptionSets = function _SpoApplicationCustomizerRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'title', 'clientSideComponentId'] });
};
export default new SpoApplicationCustomizerRemoveCommand();
//# sourceMappingURL=applicationcustomizer-remove.js.map