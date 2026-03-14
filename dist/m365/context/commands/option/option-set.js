var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _ContextOptionSetCommand_instances, _ContextOptionSetCommand_initOptions;
import fs from 'fs';
import { CommandError } from '../../../../Command.js';
import ContextCommand from '../../../base/ContextCommand.js';
import commands from '../../commands.js';
class ContextOptionSetCommand extends ContextCommand {
    get name() {
        return commands.OPTION_SET;
    }
    get description() {
        return 'Allows to add a new name for the option and value to the local context file.';
    }
    constructor() {
        super();
        _ContextOptionSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _ContextOptionSetCommand_instances, "m", _ContextOptionSetCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        const filePath = '.m365rc.json';
        if (this.verbose) {
            await logger.logToStderr(`Saving ${args.options.name} with value ${args.options.value} to the ${filePath} file...`);
        }
        let m365rc = {};
        if (fs.existsSync(filePath)) {
            try {
                if (this.verbose) {
                    await logger.logToStderr(`Reading existing ${filePath}...`);
                }
                const fileContents = fs.readFileSync(filePath, 'utf8');
                if (fileContents) {
                    m365rc = JSON.parse(fileContents);
                }
            }
            catch (e) {
                throw new CommandError(`Error reading ${filePath}: ${e}. Please add ${args.options.name} to ${filePath} manually.`);
            }
        }
        if (m365rc.context) {
            m365rc.context[args.options.name] = args.options.value;
            try {
                if (this.verbose) {
                    await logger.logToStderr(`Creating option ${args.options.name} with value ${args.options.value} in existing context...`);
                }
                fs.writeFileSync(filePath, JSON.stringify(m365rc, null, 2));
            }
            catch (e) {
                throw new CommandError(`Error writing ${filePath}: ${e}. Please add ${args.options.name} to ${filePath} manually.`);
            }
        }
        else {
            if (this.verbose) {
                await logger.logToStderr(`Context doesn't exist. Initializing the context and creating option ${args.options.name} with value ${args.options.value}...`);
            }
            this.saveContextInfo({ [args.options.name]: args.options.value });
        }
    }
}
_ContextOptionSetCommand_instances = new WeakSet(), _ContextOptionSetCommand_initOptions = function _ContextOptionSetCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name <name>'
    }, {
        option: '-v, --value <value>'
    });
};
export default new ContextOptionSetCommand();
//# sourceMappingURL=option-set.js.map