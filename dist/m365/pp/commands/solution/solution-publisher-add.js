var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PpSolutionPublisherAddCommand_instances, _PpSolutionPublisherAddCommand_initTelemetry, _PpSolutionPublisherAddCommand_initOptions, _PpSolutionPublisherAddCommand_initValidators;
import request from '../../../../request.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
class PpSolutionPublisherAddCommand extends PowerPlatformCommand {
    get name() {
        return commands.SOLUTION_PUBLISHER_ADD;
    }
    get description() {
        return 'Adds a specified publisher in a given environment';
    }
    constructor() {
        super();
        _PpSolutionPublisherAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _PpSolutionPublisherAddCommand_instances, "m", _PpSolutionPublisherAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PpSolutionPublisherAddCommand_instances, "m", _PpSolutionPublisherAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PpSolutionPublisherAddCommand_instances, "m", _PpSolutionPublisherAddCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Adding new publisher '${args.options.name}'...`);
        }
        try {
            const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);
            const requestOptions = {
                url: `${dynamicsApiUrl}/api/data/v9.0/publishers`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: {
                    uniquename: args.options.name,
                    friendlyname: args.options.displayName,
                    customizationprefix: args.options.prefix,
                    customizationoptionvalueprefix: args.options.choiceValuePrefix
                }
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_PpSolutionPublisherAddCommand_instances = new WeakSet(), _PpSolutionPublisherAddCommand_initTelemetry = function _PpSolutionPublisherAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            asAdmin: !!args.options.asAdmin
        });
    });
}, _PpSolutionPublisherAddCommand_initOptions = function _PpSolutionPublisherAddCommand_initOptions() {
    this.options.unshift({
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '-n, --name <name>'
    }, {
        option: '--displayName <displayName>'
    }, {
        option: '--prefix <prefix>'
    }, {
        option: '--choiceValuePrefix <choiceValuePrefix>'
    }, {
        option: '--asAdmin'
    });
}, _PpSolutionPublisherAddCommand_initValidators = function _PpSolutionPublisherAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (isNaN(args.options.choiceValuePrefix) || args.options.choiceValuePrefix < 10000 || args.options.choiceValuePrefix > 99999 || !Number.isInteger(args.options.choiceValuePrefix)) {
            return 'Option choiceValuePrefix should be an integer between 10000 and 99999.';
        }
        const nameRegEx = new RegExp(/^[a-zA-Z_][A-Za-z0-9_]+$/);
        if (!nameRegEx.test(args.options.name)) {
            return 'Option name may only consist of alphanumeric characters and underscores. The first character cannot be a number.';
        }
        const prefixRegEx = new RegExp(/^(?!mscrm.*$)[a-zA-Z][A-Za-z0-9]{1,7}$/);
        if (!prefixRegEx.test(args.options.prefix)) {
            return `Option prefix may only consist of alphanumeric characters. The first character cannot be a number and may not start with 'mscrm'. It must be between 2 and 8 characters long.`;
        }
        return true;
    });
};
export default new PpSolutionPublisherAddCommand();
//# sourceMappingURL=solution-publisher-add.js.map