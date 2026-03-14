var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PaAppConsentSetCommand_instances, _PaAppConsentSetCommand_initOptions, _PaAppConsentSetCommand_initValidators, _PaAppConsentSetCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import PowerAppsCommand from '../../../base/PowerAppsCommand.js';
import commands from '../../commands.js';
class PaAppConsentSetCommand extends PowerAppsCommand {
    get name() {
        return commands.APP_CONSENT_SET;
    }
    get description() {
        return 'Configures if users can bypass the API Consent window for the selected canvas app';
    }
    constructor() {
        super();
        _PaAppConsentSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _PaAppConsentSetCommand_instances, "m", _PaAppConsentSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PaAppConsentSetCommand_instances, "m", _PaAppConsentSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PaAppConsentSetCommand_instances, "m", _PaAppConsentSetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Setting the bypass consent for the Microsoft Power App ${args.options.name}... to ${args.options.bypass}`);
        }
        if (args.options.force) {
            await this.consentPaApp(args);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you bypass the consent for the Microsoft Power App ${args.options.name} to ${args.options.bypass}?` });
            if (result) {
                await this.consentPaApp(args);
            }
        }
    }
    async consentPaApp(args) {
        const requestOptions = {
            url: `${this.resource}/providers/Microsoft.PowerApps/scopes/admin/environments/${args.options.environmentName}/apps/${args.options.name}/setPowerAppConnectionDirectConsentBypass?api-version=2021-02-01`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json',
            data: {
                bypassconsent: args.options.bypass
            }
        };
        try {
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_PaAppConsentSetCommand_instances = new WeakSet(), _PaAppConsentSetCommand_initOptions = function _PaAppConsentSetCommand_initOptions() {
    this.options.unshift({
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '-n, --name <name>'
    }, {
        option: '-b, --bypass <bypass>',
        autocomplete: ['true', 'false']
    }, {
        option: '-f, --force'
    });
}, _PaAppConsentSetCommand_initValidators = function _PaAppConsentSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.name)) {
            return `${args.options.name} is not a valid GUID`;
        }
        return true;
    });
}, _PaAppConsentSetCommand_initTypes = function _PaAppConsentSetCommand_initTypes() {
    this.types.boolean.push('bypass');
};
export default new PaAppConsentSetCommand();
//# sourceMappingURL=app-consent-set.js.map