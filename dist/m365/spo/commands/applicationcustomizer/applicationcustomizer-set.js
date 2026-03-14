var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoApplicationCustomizerSetCommand_instances, _SpoApplicationCustomizerSetCommand_initOptions, _SpoApplicationCustomizerSetCommand_initTelemetry, _SpoApplicationCustomizerSetCommand_initValidators, _SpoApplicationCustomizerSetCommand_initOptionSets;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { spo } from '../../../../utils/spo.js';
import { formatting } from '../../../../utils/formatting.js';
import { cli } from '../../../../cli/cli.js';
class SpoApplicationCustomizerSetCommand extends SpoCommand {
    get name() {
        return commands.APPLICATIONCUSTOMIZER_SET;
    }
    get description() {
        return 'Updates an existing Application Customizer on a site';
    }
    constructor() {
        super();
        _SpoApplicationCustomizerSetCommand_instances.add(this);
        this.allowedScopes = ['Site', 'Web', 'All'];
        __classPrivateFieldGet(this, _SpoApplicationCustomizerSetCommand_instances, "m", _SpoApplicationCustomizerSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoApplicationCustomizerSetCommand_instances, "m", _SpoApplicationCustomizerSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoApplicationCustomizerSetCommand_instances, "m", _SpoApplicationCustomizerSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoApplicationCustomizerSetCommand_instances, "m", _SpoApplicationCustomizerSetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const appCustomizer = await this.getAppCustomizerToUpdate(logger, args.options);
            await this.updateAppCustomizer(logger, args.options, appCustomizer);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async updateAppCustomizer(logger, options, appCustomizer) {
        const { clientSideComponentProperties, webUrl, newTitle, description } = options;
        if (this.verbose) {
            await logger.logToStderr(`Updating application customizer with ID '${appCustomizer.Id}' on the site '${webUrl}'...`);
        }
        const requestBody = {};
        if (newTitle) {
            requestBody.Title = newTitle;
        }
        if (description !== undefined) {
            requestBody.Description = description;
        }
        if (clientSideComponentProperties !== undefined) {
            requestBody.ClientSideComponentProperties = clientSideComponentProperties;
        }
        const requestOptions = {
            url: `${webUrl}/_api/${appCustomizer.Scope.toString() === '2' ? 'Site' : 'Web'}/UserCustomActions('${appCustomizer.Id}')`,
            headers: {
                accept: 'application/json;odata=nometadata',
                'X-HTTP-Method': 'MERGE'
            },
            data: requestBody,
            responseType: 'json'
        };
        await request.post(requestOptions);
    }
    async getAppCustomizerToUpdate(logger, options) {
        const { id, webUrl, title, clientSideComponentId, scope } = options;
        const resolvedScope = scope || 'All';
        if (this.verbose) {
            await logger.logToStderr(`Getting application customizer ${title || clientSideComponentId || id} to update...`);
        }
        let appCustomizers = [];
        if (id) {
            const appCustomizer = await spo.getCustomActionById(webUrl, id, resolvedScope);
            if (appCustomizer && appCustomizer.Location === 'ClientSideExtension.ApplicationCustomizer') {
                appCustomizers.push(appCustomizer);
            }
        }
        else if (title) {
            appCustomizers = await spo.getCustomActions(webUrl, resolvedScope, `(Title eq '${formatting.encodeQueryParameter(title)}') and (startswith(Location,'ClientSideExtension.ApplicationCustomizer'))`);
        }
        else {
            appCustomizers = await spo.getCustomActions(webUrl, resolvedScope, `(ClientSideComponentId eq guid'${clientSideComponentId}') and (startswith(Location,'ClientSideExtension.ApplicationCustomizer'))`);
        }
        if (appCustomizers.length === 0) {
            throw `No application customizer with ${title && `title '${title}'` || clientSideComponentId && `ClientSideComponentId '${clientSideComponentId}'` || id && `id '${id}'`} found`;
        }
        if (appCustomizers.length > 1) {
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', appCustomizers);
            return await cli.handleMultipleResultsFound(`Multiple application customizer with ${title ? `title '${title}'` : `ClientSideComponentId '${clientSideComponentId}'`} found.`, resultAsKeyValuePair);
        }
        return appCustomizers[0];
    }
}
_SpoApplicationCustomizerSetCommand_instances = new WeakSet(), _SpoApplicationCustomizerSetCommand_initOptions = function _SpoApplicationCustomizerSetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-t, --title [title]'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-c, --clientSideComponentId [clientSideComponentId]'
    }, {
        option: '--newTitle [newTitle]'
    }, {
        option: '--description [description]'
    }, {
        option: '-p, --clientSideComponentProperties [clientSideComponentProperties]'
    }, {
        option: '-s, --scope [scope]', autocomplete: this.allowedScopes
    });
}, _SpoApplicationCustomizerSetCommand_initTelemetry = function _SpoApplicationCustomizerSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            title: typeof args.options.title !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            clientSideComponentId: typeof args.options.clientSideComponentId !== 'undefined',
            newTitle: typeof args.options.newTitle !== 'undefined',
            description: typeof args.options.description !== 'undefined',
            clientSideComponentProperties: typeof args.options.clientSideComponentProperties !== 'undefined',
            scope: typeof args.options.scope !== 'undefined'
        });
    });
}, _SpoApplicationCustomizerSetCommand_initValidators = function _SpoApplicationCustomizerSetCommand_initValidators() {
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
        if (!args.options.newTitle && args.options.description === undefined && !args.options.clientSideComponentProperties) {
            return `Please specify an option to be updated`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
}, _SpoApplicationCustomizerSetCommand_initOptionSets = function _SpoApplicationCustomizerSetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'title', 'clientSideComponentId'] });
};
export default new SpoApplicationCustomizerSetCommand();
//# sourceMappingURL=applicationcustomizer-set.js.map