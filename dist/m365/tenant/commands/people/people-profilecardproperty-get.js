var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TenantPeopleProfileCardPropertyGetCommand_instances, _TenantPeopleProfileCardPropertyGetCommand_initOptions, _TenantPeopleProfileCardPropertyGetCommand_initValidators;
import GraphCommand from '../../../base/GraphCommand.js';
import request from '../../../../request.js';
import { profileCardPropertyNames } from './profileCardProperties.js';
import commands from '../../commands.js';
class TenantPeopleProfileCardPropertyGetCommand extends GraphCommand {
    get name() {
        return commands.PEOPLE_PROFILECARDPROPERTY_GET;
    }
    get description() {
        return 'Retrieves information about a specific profile card property';
    }
    constructor() {
        super();
        _TenantPeopleProfileCardPropertyGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _TenantPeopleProfileCardPropertyGetCommand_instances, "m", _TenantPeopleProfileCardPropertyGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TenantPeopleProfileCardPropertyGetCommand_instances, "m", _TenantPeopleProfileCardPropertyGetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Retrieving information about profile card property '${args.options.name}'...`);
            }
            // Get the right casing for the profile card property name
            const profileCardProperty = profileCardPropertyNames.find(p => p.toLowerCase() === args.options.name.toLowerCase());
            const requestOptions = {
                url: `${this.resource}/v1.0/admin/people/profileCardProperties/${profileCardProperty}`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            const result = await request.get(requestOptions);
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
            if (err.response?.status === 404) {
                this.handleError(`Profile card property '${args.options.name}' does not exist.`);
            }
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_TenantPeopleProfileCardPropertyGetCommand_instances = new WeakSet(), _TenantPeopleProfileCardPropertyGetCommand_initOptions = function _TenantPeopleProfileCardPropertyGetCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name <name>',
        autocomplete: profileCardPropertyNames
    });
}, _TenantPeopleProfileCardPropertyGetCommand_initValidators = function _TenantPeopleProfileCardPropertyGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!profileCardPropertyNames.some(p => p.toLowerCase() === args.options.name.toLowerCase())) {
            return `'${args.options.name}' is not a valid value for option name. Allowed values are: ${profileCardPropertyNames.join(', ')}.`;
        }
        return true;
    });
};
export default new TenantPeopleProfileCardPropertyGetCommand();
//# sourceMappingURL=people-profilecardproperty-get.js.map