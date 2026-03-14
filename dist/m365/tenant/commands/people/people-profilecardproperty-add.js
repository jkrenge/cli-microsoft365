var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TenantPeopleProfileCardPropertyAddCommand_instances, _TenantPeopleProfileCardPropertyAddCommand_initTelemetry, _TenantPeopleProfileCardPropertyAddCommand_initOptions, _TenantPeopleProfileCardPropertyAddCommand_initValidators;
import request from '../../../../request.js';
import { optionsUtils } from '../../../../utils/optionsUtils.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { profileCardPropertyNames } from './profileCardProperties.js';
class TenantPeopleProfileCardPropertyAddCommand extends GraphCommand {
    get name() {
        return commands.PEOPLE_PROFILECARDPROPERTY_ADD;
    }
    get description() {
        return 'Adds an additional attribute to the profile card properties';
    }
    constructor() {
        super();
        _TenantPeopleProfileCardPropertyAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _TenantPeopleProfileCardPropertyAddCommand_instances, "m", _TenantPeopleProfileCardPropertyAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TenantPeopleProfileCardPropertyAddCommand_instances, "m", _TenantPeopleProfileCardPropertyAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TenantPeopleProfileCardPropertyAddCommand_instances, "m", _TenantPeopleProfileCardPropertyAddCommand_initValidators).call(this);
    }
    allowUnknownOptions() {
        return true;
    }
    async commandAction(logger, args) {
        const directoryPropertyName = profileCardPropertyNames.find(n => n.toLowerCase() === args.options.name.toLowerCase());
        if (this.verbose) {
            await logger.logToStderr(`Adding '${directoryPropertyName}' as a profile card property...`);
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/admin/people/profileCardProperties`,
            headers: {
                'content-type': 'application/json',
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json',
            data: {
                directoryPropertyName,
                annotations: this.getAnnotations(args.options)
            }
        };
        try {
            const response = await request.post(requestOptions);
            // Transform the output to make it more readable
            if (args.options.output && args.options.output !== 'json' && response.annotations.length > 0) {
                const annotation = response.annotations[0];
                response.displayName = annotation.displayName;
                annotation.localizations.forEach((l) => {
                    response[`displayName ${l.languageTag}`] = l.displayName;
                });
                delete response.annotations;
            }
            await logger.log(response);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getAnnotations(options) {
        if (!options.displayName) {
            return [];
        }
        return [
            {
                displayName: options.displayName,
                localizations: this.getLocalizations(options)
            }
        ];
    }
    getLocalizations(options) {
        const unknownOptions = Object.keys(optionsUtils.getUnknownOptions(options, this.options));
        if (unknownOptions.length === 0) {
            return [];
        }
        const localizations = [];
        unknownOptions.forEach(key => {
            localizations.push({
                languageTag: key.replace('displayName-', ''),
                displayName: options[key]
            });
        });
        return localizations;
    }
}
_TenantPeopleProfileCardPropertyAddCommand_instances = new WeakSet(), _TenantPeopleProfileCardPropertyAddCommand_initTelemetry = function _TenantPeopleProfileCardPropertyAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        // Add unknown options to telemetry
        const unknownOptions = Object.keys(optionsUtils.getUnknownOptions(args.options, this.options));
        const unknownOptionsObj = unknownOptions.reduce((obj, key) => ({ ...obj, [key]: true }), {});
        Object.assign(this.telemetryProperties, {
            displayName: typeof args.options.displayName !== 'undefined',
            ...unknownOptionsObj
        });
    });
}, _TenantPeopleProfileCardPropertyAddCommand_initOptions = function _TenantPeopleProfileCardPropertyAddCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name <name>',
        autocomplete: profileCardPropertyNames
    }, {
        option: '-d, --displayName [displayName]'
    });
}, _TenantPeopleProfileCardPropertyAddCommand_initValidators = function _TenantPeopleProfileCardPropertyAddCommand_initValidators() {
    this.validators.push(async (args) => {
        const propertyName = args.options.name.toLowerCase();
        if (profileCardPropertyNames.every(n => n.toLowerCase() !== propertyName)) {
            return `${args.options.name} is not a valid value for name. Allowed values are ${profileCardPropertyNames.join(', ')}`;
        }
        if (propertyName.startsWith('customattribute') && args.options.displayName === undefined) {
            return `The option 'displayName' is required when adding customAttributes as profile card properties`;
        }
        if (!propertyName.startsWith('customattribute') && args.options.displayName !== undefined) {
            return `The option 'displayName' can only be used when adding customAttributes as profile card properties`;
        }
        const unknownOptions = Object.keys(optionsUtils.getUnknownOptions(args.options, this.options));
        if (!propertyName.startsWith('customattribute') && unknownOptions.length > 0) {
            return `Unknown options like ${unknownOptions.join(', ')} are only supported with customAttributes`;
        }
        if (propertyName.startsWith('customattribute')) {
            const wronglyFormattedOptions = unknownOptions.filter(key => !key.toLowerCase().startsWith('displayname-'));
            if (wronglyFormattedOptions.length > 0) {
                return `Wrong option format detected for the following option(s): ${wronglyFormattedOptions.join(', ')}'. When adding localizations for customAttributes, use the format displayName-<languageTag>.`;
            }
        }
        return true;
    });
};
export default new TenantPeopleProfileCardPropertyAddCommand();
//# sourceMappingURL=people-profilecardproperty-add.js.map