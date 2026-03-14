var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _CliCompletionPwshSetupCommand_instances, _CliCompletionPwshSetupCommand_initOptions;
import fs from 'fs';
import os from 'os';
import path from 'path';
import url from 'url';
import { autocomplete } from '../../../../autocomplete.js';
import { CommandError } from '../../../../Command.js';
import AnonymousCommand from '../../../base/AnonymousCommand.js';
import commands from '../../commands.js';
const __dirname = url.fileURLToPath(new URL('.', import.meta.url));
class CliCompletionPwshSetupCommand extends AnonymousCommand {
    get name() {
        return commands.COMPLETION_PWSH_SETUP;
    }
    get description() {
        return 'Sets up command completion for PowerShell';
    }
    constructor() {
        super();
        _CliCompletionPwshSetupCommand_instances.add(this);
        __classPrivateFieldGet(this, _CliCompletionPwshSetupCommand_instances, "m", _CliCompletionPwshSetupCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        if (this.debug) {
            await logger.logToStderr('Generating command completion...');
        }
        autocomplete.generateShCompletion();
        if (this.debug) {
            await logger.logToStderr(`Ensuring that the specified profile path ${args.options.profile} exists...`);
        }
        if (fs.existsSync(args.options.profile)) {
            if (this.debug) {
                await logger.logToStderr('Profile file already exists');
            }
        }
        else {
            // check if the path exists
            const dirname = path.dirname(args.options.profile);
            if (fs.existsSync(dirname)) {
                if (this.debug) {
                    await logger.logToStderr(`Profile path ${dirname} already exists`);
                }
            }
            else {
                try {
                    if (this.debug) {
                        await logger.logToStderr(`Profile path ${dirname} doesn't exist. Creating...`);
                    }
                    fs.mkdirSync(dirname, { recursive: true });
                }
                catch (e) {
                    throw new CommandError(e);
                }
            }
            if (this.debug) {
                await logger.logToStderr(`Creating profile file ${args.options.profile}...`);
            }
            try {
                fs.writeFileSync(args.options.profile, '', 'utf8');
            }
            catch (e) {
                throw new CommandError(e);
            }
        }
        if (this.verbose) {
            await logger.logToStderr(`Adding CLI for Microsoft 365 command completion to PowerShell profile...`);
        }
        const completionScriptPath = path.resolve(__dirname, '..', '..', '..', '..', '..', 'scripts', 'Register-CLIM365Completion.ps1');
        try {
            fs.appendFileSync(args.options.profile, os.EOL + completionScriptPath, 'utf8');
            return;
        }
        catch (e) {
            throw new CommandError(e);
        }
    }
}
_CliCompletionPwshSetupCommand_instances = new WeakSet(), _CliCompletionPwshSetupCommand_initOptions = function _CliCompletionPwshSetupCommand_initOptions() {
    this.options.unshift({
        option: '-p, --profile <profile>'
    });
};
export default new CliCompletionPwshSetupCommand();
//# sourceMappingURL=completion-pwsh-setup.js.map