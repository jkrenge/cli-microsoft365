var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoWebRoleInheritanceBreakCommand_instances, _SpoWebRoleInheritanceBreakCommand_initTelemetry, _SpoWebRoleInheritanceBreakCommand_initOptions, _SpoWebRoleInheritanceBreakCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoWebRoleInheritanceBreakCommand extends SpoCommand {
    get name() {
        return commands.WEB_ROLEINHERITANCE_BREAK;
    }
    get description() {
        return 'Break role inheritance of subsite';
    }
    constructor() {
        super();
        _SpoWebRoleInheritanceBreakCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoWebRoleInheritanceBreakCommand_instances, "m", _SpoWebRoleInheritanceBreakCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoWebRoleInheritanceBreakCommand_instances, "m", _SpoWebRoleInheritanceBreakCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoWebRoleInheritanceBreakCommand_instances, "m", _SpoWebRoleInheritanceBreakCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Break role inheritance of subsite with URL ${args.options.webUrl}...`);
        }
        if (args.options.force) {
            await this.breakRoleInheritance(args.options);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to break the role inheritance of subsite ${args.options.webUrl}?` });
            if (result) {
                await this.breakRoleInheritance(args.options);
            }
        }
    }
    async breakRoleInheritance(options) {
        const requestOptions = {
            url: `${options.webUrl}/_api/web/breakroleinheritance(${!options.clearExistingPermissions})`,
            headers: {
                'accept': 'application/json;odata=nometadata',
                'content-type': 'application/json'
            },
            responseType: 'json'
        };
        try {
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoWebRoleInheritanceBreakCommand_instances = new WeakSet(), _SpoWebRoleInheritanceBreakCommand_initTelemetry = function _SpoWebRoleInheritanceBreakCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            clearExistingPermissions: !!args.options.clearExistingPermissions,
            force: !!args.options.force
        });
    });
}, _SpoWebRoleInheritanceBreakCommand_initOptions = function _SpoWebRoleInheritanceBreakCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-c, --clearExistingPermissions'
    }, {
        option: '-f, --force'
    });
}, _SpoWebRoleInheritanceBreakCommand_initValidators = function _SpoWebRoleInheritanceBreakCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        return true;
    });
};
export default new SpoWebRoleInheritanceBreakCommand();
//# sourceMappingURL=web-roleinheritance-break.js.map