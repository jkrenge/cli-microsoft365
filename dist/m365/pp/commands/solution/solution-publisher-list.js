var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PpSolutionPublisherListCommand_instances, _PpSolutionPublisherListCommand_initTelemetry, _PpSolutionPublisherListCommand_initOptions;
import request from '../../../../request.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
class PpSolutionPublisherListCommand extends PowerPlatformCommand {
    get name() {
        return commands.SOLUTION_PUBLISHER_LIST;
    }
    get description() {
        return 'Lists publishers in a given environment.';
    }
    defaultProperties() {
        return ['publisherid', 'uniquename', 'friendlyname'];
    }
    constructor() {
        super();
        _PpSolutionPublisherListCommand_instances.add(this);
        __classPrivateFieldGet(this, _PpSolutionPublisherListCommand_instances, "m", _PpSolutionPublisherListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PpSolutionPublisherListCommand_instances, "m", _PpSolutionPublisherListCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving list of publishers...`);
        }
        try {
            const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);
            const requestOptions = {
                url: `${dynamicsApiUrl}/api/data/v9.0/publishers?$select=publisherid,uniquename,friendlyname,versionnumber,isreadonly,description,customizationprefix,customizationoptionvalueprefix${!args.options.withMicrosoftPublishers ? `&$filter=publisherid ne 'd21aab70-79e7-11dd-8874-00188b01e34f'` : ''}&api-version=9.1`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            const res = await request.get(requestOptions);
            await logger.log(res.value);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_PpSolutionPublisherListCommand_instances = new WeakSet(), _PpSolutionPublisherListCommand_initTelemetry = function _PpSolutionPublisherListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            withMicrosoftPublishers: typeof args.options.withMicrosoftPublishers !== 'undefined',
            asAdmin: !!args.options.asAdmin
        });
    });
}, _PpSolutionPublisherListCommand_initOptions = function _PpSolutionPublisherListCommand_initOptions() {
    this.options.unshift({
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '--withMicrosoftPublishers'
    }, {
        option: '--asAdmin'
    });
};
export default new PpSolutionPublisherListCommand();
//# sourceMappingURL=solution-publisher-list.js.map