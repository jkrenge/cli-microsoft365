var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoHubSiteUnregisterCommand_instances, _SpoHubSiteUnregisterCommand_initTelemetry, _SpoHubSiteUnregisterCommand_initOptions, _SpoHubSiteUnregisterCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoHubSiteUnregisterCommand extends SpoCommand {
    get name() {
        return commands.HUBSITE_UNREGISTER;
    }
    get description() {
        return 'Unregisters the specified site collection as a hub site';
    }
    constructor() {
        super();
        _SpoHubSiteUnregisterCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoHubSiteUnregisterCommand_instances, "m", _SpoHubSiteUnregisterCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoHubSiteUnregisterCommand_instances, "m", _SpoHubSiteUnregisterCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoHubSiteUnregisterCommand_instances, "m", _SpoHubSiteUnregisterCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const unregisterHubSite = async () => {
            try {
                if (this.verbose) {
                    await logger.logToStderr(`Unregistering site collection ${args.options.url} as a hub site...`);
                }
                const res = await spo.getRequestDigest(args.options.url);
                const requestOptions = {
                    url: `${args.options.url}/_api/site/UnregisterHubSite`,
                    headers: {
                        'X-RequestDigest': res.FormDigestValue,
                        accept: 'application/json;odata=nometadata'
                    },
                    responseType: 'json'
                };
                await request.post(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await unregisterHubSite();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to unregister the site collection ${args.options.url} as a hub site?` });
            if (result) {
                await unregisterHubSite();
            }
        }
    }
}
_SpoHubSiteUnregisterCommand_instances = new WeakSet(), _SpoHubSiteUnregisterCommand_initTelemetry = function _SpoHubSiteUnregisterCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: args.options.force || false
        });
    });
}, _SpoHubSiteUnregisterCommand_initOptions = function _SpoHubSiteUnregisterCommand_initOptions() {
    this.options.unshift({
        option: '-u, --url <url>'
    }, {
        option: '-f, --force'
    });
}, _SpoHubSiteUnregisterCommand_initValidators = function _SpoHubSiteUnregisterCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.url));
};
export default new SpoHubSiteUnregisterCommand();
//# sourceMappingURL=hubsite-unregister.js.map