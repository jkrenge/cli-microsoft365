var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoHubSiteDataGetCommand_instances, _SpoHubSiteDataGetCommand_initTelemetry, _SpoHubSiteDataGetCommand_initOptions, _SpoHubSiteDataGetCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoHubSiteDataGetCommand extends SpoCommand {
    get name() {
        return commands.HUBSITE_DATA_GET;
    }
    get description() {
        return 'Get hub site data for the specified site';
    }
    constructor() {
        super();
        _SpoHubSiteDataGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoHubSiteDataGetCommand_instances, "m", _SpoHubSiteDataGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoHubSiteDataGetCommand_instances, "m", _SpoHubSiteDataGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoHubSiteDataGetCommand_instances, "m", _SpoHubSiteDataGetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr('Retrieving hub site data...');
        }
        const forceRefresh = args.options.forceRefresh === true;
        const requestOptions = {
            url: `${args.options.webUrl}/_api/web/HubSiteData(${forceRefresh})`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        try {
            const res = await request.get(requestOptions);
            if (res['odata.null'] !== true) {
                await logger.log(JSON.parse(res.value));
            }
            else {
                if (this.verbose) {
                    await logger.logToStderr(`${args.options.webUrl} is not connected to a hub site and is not a hub site itself`);
                }
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoHubSiteDataGetCommand_instances = new WeakSet(), _SpoHubSiteDataGetCommand_initTelemetry = function _SpoHubSiteDataGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            forceRefresh: args.options.forceRefresh === true
        });
    });
}, _SpoHubSiteDataGetCommand_initOptions = function _SpoHubSiteDataGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--forceRefresh'
    });
}, _SpoHubSiteDataGetCommand_initValidators = function _SpoHubSiteDataGetCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.webUrl));
};
export default new SpoHubSiteDataGetCommand();
//# sourceMappingURL=hubsite-data-get.js.map