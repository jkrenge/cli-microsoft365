var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoNavigationNodeRemoveCommand_instances, _SpoNavigationNodeRemoveCommand_initTelemetry, _SpoNavigationNodeRemoveCommand_initOptions, _SpoNavigationNodeRemoveCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoNavigationNodeRemoveCommand extends SpoCommand {
    get name() {
        return commands.NAVIGATION_NODE_REMOVE;
    }
    get description() {
        return 'Removes the specified navigation node';
    }
    constructor() {
        super();
        _SpoNavigationNodeRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoNavigationNodeRemoveCommand_instances, "m", _SpoNavigationNodeRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoNavigationNodeRemoveCommand_instances, "m", _SpoNavigationNodeRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoNavigationNodeRemoveCommand_instances, "m", _SpoNavigationNodeRemoveCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeNode(logger, args.options);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the node ${args.options.id} from the navigation?` });
            if (result) {
                await this.removeNode(logger, args.options);
            }
        }
    }
    async removeNode(logger, options) {
        try {
            const res = await spo.getRequestDigest(options.webUrl);
            if (this.verbose) {
                await logger.logToStderr(`Removing navigation node...`);
            }
            const requestOptions = {
                url: `${options.webUrl}/_api/web/navigation/${options.location.toLowerCase()}/getbyid(${options.id})`,
                headers: {
                    accept: 'application/json;odata=nometadata',
                    'X-RequestDigest': res.FormDigestValue
                },
                responseType: 'json'
            };
            await request.delete(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoNavigationNodeRemoveCommand_instances = new WeakSet(), _SpoNavigationNodeRemoveCommand_initTelemetry = function _SpoNavigationNodeRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            location: args.options.location,
            force: typeof args.options.force !== 'undefined'
        });
    });
}, _SpoNavigationNodeRemoveCommand_initOptions = function _SpoNavigationNodeRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-l, --location <location>',
        autocomplete: ['QuickLaunch', 'TopNavigationBar']
    }, {
        option: '-i, --id <id>'
    }, {
        option: '-f, --force'
    });
}, _SpoNavigationNodeRemoveCommand_initValidators = function _SpoNavigationNodeRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.location !== 'QuickLaunch' &&
            args.options.location !== 'TopNavigationBar') {
            return `${args.options.location} is not a valid value for the location option. Allowed values are QuickLaunch|TopNavigationBar`;
        }
        const id = parseInt(args.options.id);
        if (isNaN(id)) {
            return `${args.options.id} is not a number`;
        }
        return true;
    });
};
export default new SpoNavigationNodeRemoveCommand();
//# sourceMappingURL=navigation-node-remove.js.map