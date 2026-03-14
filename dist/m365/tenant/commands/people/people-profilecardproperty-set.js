var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TenantPeopleProfileCardPropertySetCommand_instances, _TenantPeopleProfileCardPropertySetCommand_initTelemetry, _TenantPeopleProfileCardPropertySetCommand_initOptions, _TenantPeopleProfileCardPropertySetCommand_initValidators, _TenantPeopleProfileCardPropertySetCommand_initTypes;
import GraphCommand from '../../../base/GraphCommand.js';
import request from '../../../../request.js';
import { profileCardPropertyNames as allProfileCardPropertyNames } from './profileCardProperties.js';
import commands from '../../commands.js';
import { optionsUtils } from '../../../../utils/optionsUtils.js';
class TenantPeopleProfileCardPropertySetCommand extends GraphCommand {
    get name() {
        return commands.PEOPLE_PROFILECARDPROPERTY_SET;
    }
    get description() {
        return 'Updates a custom attribute to the profile card property';
    }
    constructor() {
        super();
        _TenantPeopleProfileCardPropertySetCommand_instances.add(this);
        this.profileCardPropertyNames = allProfileCardPropertyNames.filter(p => p.toLowerCase().startsWith('customattribute'));
        __classPrivateFieldGet(this, _TenantPeopleProfileCardPropertySetCommand_instances, "m", _TenantPeopleProfileCardPropertySetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TenantPeopleProfileCardPropertySetCommand_instances, "m", _TenantPeopleProfileCardPropertySetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TenantPeopleProfileCardPropertySetCommand_instances, "m", _TenantPeopleProfileCardPropertySetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TenantPeopleProfileCardPropertySetCommand_instances, "m", _TenantPeopleProfileCardPropertySetCommand_initTypes).call(this);
    }
    allowUnknownOptions() {
        return true;
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Updating profile card property '${args.options.name}'...`);
            }
            // Get the right casing for the profile card property name
            const profileCardProperty = this.profileCardPropertyNames.find(p => p.toLowerCase() === args.options.name.toLowerCase());
            const requestOptions = {
                url: `${this.resource}/v1.0/admin/people/profileCardProperties/${profileCardProperty}`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: {
                    annotations: [
                        {
                            displayName: args.options.displayName,
                            localizations: this.getLocalizations(args.options)
                        }
                    ]
                }
            };
            const result = await request.patch(requestOptions);
            let output = result;
            // Transform the output to make it more readable
            if (args.options.output && args.options.output !== 'json' && result.annotations.length > 0) {
                output = result.annotations[0].localizations.reduce((acc, curr) => ({
                    ...acc,
                    ['displayName ' + curr.languageTag]: curr.displayName
                }), {
                    ...result,
                    displayName: result.annotations[0].displayName
                });
                delete output.annotations;
            }
            await logger.log(output);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    /**
     * Transform option to localization object.
     * @example Transform "--displayName-en-US 'Cost center'" to { languageTag: 'en-US', displayName: 'Cost center' }
     */
    getLocalizations(options) {
        const unknownOptions = optionsUtils.getUnknownOptions(options, this.options);
        const result = Object.keys(unknownOptions).map(o => ({
            languageTag: o.substring(o.indexOf('-') + 1),
            displayName: unknownOptions[o]
        }));
        return result;
    }
}
_TenantPeopleProfileCardPropertySetCommand_instances = new WeakSet(), _TenantPeopleProfileCardPropertySetCommand_initTelemetry = function _TenantPeopleProfileCardPropertySetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        // Add unknown options to telemetry
        const unknownOptions = Object.keys(optionsUtils.getUnknownOptions(args.options, this.options));
        const unknownOptionsObj = unknownOptions.reduce((obj, key) => ({ ...obj, [key]: true }), {});
        Object.assign(this.telemetryProperties, {
            displayName: typeof args.options.displayName !== 'undefined',
            ...unknownOptionsObj
        });
    });
}, _TenantPeopleProfileCardPropertySetCommand_initOptions = function _TenantPeopleProfileCardPropertySetCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name <name>',
        autocomplete: this.profileCardPropertyNames
    }, {
        option: '-d, --displayName <displayName>'
    });
}, _TenantPeopleProfileCardPropertySetCommand_initValidators = function _TenantPeopleProfileCardPropertySetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!this.profileCardPropertyNames.some(p => p.toLowerCase() === args.options.name.toLowerCase())) {
            return `'${args.options.name}' is not a valid value for option name. Allowed values are: ${this.profileCardPropertyNames.join(', ')}.`;
        }
        // Unknown options are allowed only if they start with 'displayName-'
        const unknownOptionKeys = Object.keys(optionsUtils.getUnknownOptions(args.options, this.options));
        const invalidOptionKey = unknownOptionKeys.find(o => !o.startsWith('displayName-'));
        if (invalidOptionKey) {
            return `Invalid option: '${invalidOptionKey}'`;
        }
        return true;
    });
}, _TenantPeopleProfileCardPropertySetCommand_initTypes = function _TenantPeopleProfileCardPropertySetCommand_initTypes() {
    this.types.string.push('name', 'displayName');
};
export default new TenantPeopleProfileCardPropertySetCommand();
//# sourceMappingURL=people-profilecardproperty-set.js.map