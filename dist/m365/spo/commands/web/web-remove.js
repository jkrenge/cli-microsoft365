var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoWebRemoveCommand_instances, _SpoWebRemoveCommand_initTelemetry, _SpoWebRemoveCommand_initOptions, _SpoWebRemoveCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoWebRemoveCommand extends SpoCommand {
    get name() {
        return commands.WEB_REMOVE;
    }
    get description() {
        return 'Delete specified subsite';
    }
    constructor() {
        super();
        _SpoWebRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoWebRemoveCommand_instances, "m", _SpoWebRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoWebRemoveCommand_instances, "m", _SpoWebRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoWebRemoveCommand_instances, "m", _SpoWebRemoveCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeWeb(logger, args.options.url);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the subsite ${args.options.url}?` });
            if (result) {
                await this.removeWeb(logger, args.options.url);
            }
        }
    }
    async removeWeb(logger, url) {
        const requestOptions = {
            url: `${encodeURI(url)}/_api/web`,
            headers: {
                accept: 'application/json;odata=nometadata',
                'X-HTTP-Method': 'DELETE'
            },
            responseType: 'json'
        };
        if (this.verbose) {
            await logger.logToStderr(`Deleting subsite ${url} ...`);
        }
        try {
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoWebRemoveCommand_instances = new WeakSet(), _SpoWebRemoveCommand_initTelemetry = function _SpoWebRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: (!(!args.options.force)).toString()
        });
    });
}, _SpoWebRemoveCommand_initOptions = function _SpoWebRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --url <url>'
    }, {
        option: '-f, --force'
    });
}, _SpoWebRemoveCommand_initValidators = function _SpoWebRemoveCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.url));
};
export default new SpoWebRemoveCommand();
//# sourceMappingURL=web-remove.js.map