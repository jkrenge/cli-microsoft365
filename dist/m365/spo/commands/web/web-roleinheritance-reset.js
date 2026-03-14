var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoWebRoleInheritanceResetCommand_instances, _SpoWebRoleInheritanceResetCommand_initTelemetry, _SpoWebRoleInheritanceResetCommand_initOptions, _SpoWebRoleInheritanceResetCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoWebRoleInheritanceResetCommand extends SpoCommand {
    get name() {
        return commands.WEB_ROLEINHERITANCE_RESET;
    }
    get description() {
        return 'Restores role inheritance of subsite';
    }
    constructor() {
        super();
        _SpoWebRoleInheritanceResetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoWebRoleInheritanceResetCommand_instances, "m", _SpoWebRoleInheritanceResetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoWebRoleInheritanceResetCommand_instances, "m", _SpoWebRoleInheritanceResetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoWebRoleInheritanceResetCommand_instances, "m", _SpoWebRoleInheritanceResetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Restore role inheritance of subsite at ${args.options.webUrl}...`);
        }
        if (args.options.force) {
            await this.resetWebRoleInheritance(args.options);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to reset the role inheritance of ${args.options.webUrl}?` });
            if (result) {
                await this.resetWebRoleInheritance(args.options);
            }
        }
    }
    async resetWebRoleInheritance(options) {
        try {
            const requestOptions = {
                url: `${options.webUrl}/_api/web/resetroleinheritance`,
                method: 'POST',
                headers: {
                    'accept': 'application/json;odata=nometadata',
                    'content-type': 'application/json'
                },
                responseType: 'json'
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoWebRoleInheritanceResetCommand_instances = new WeakSet(), _SpoWebRoleInheritanceResetCommand_initTelemetry = function _SpoWebRoleInheritanceResetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: (!(!args.options.force)).toString()
        });
    });
}, _SpoWebRoleInheritanceResetCommand_initOptions = function _SpoWebRoleInheritanceResetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-f, --force'
    });
}, _SpoWebRoleInheritanceResetCommand_initValidators = function _SpoWebRoleInheritanceResetCommand_initValidators() {
    this.validators.push(async (args) => {
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
};
export default new SpoWebRoleInheritanceResetCommand();
//# sourceMappingURL=web-roleinheritance-reset.js.map