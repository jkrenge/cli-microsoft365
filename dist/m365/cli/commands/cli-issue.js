var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _CliIssueCommand_instances, _a, _CliIssueCommand_initTelemetry, _CliIssueCommand_initOptions, _CliIssueCommand_initValidators;
import { browserUtil } from '../../../utils/browserUtil.js';
import AnonymousCommand from '../../base/AnonymousCommand.js';
import commands from '../commands.js';
class CliIssueCommand extends AnonymousCommand {
    get name() {
        return commands.ISSUE;
    }
    get description() {
        return 'Returns, or opens a URL that takes the user to the right place in the CLI GitHub repo to create a new issue reporting bug, feedback, ideas, etc.';
    }
    constructor() {
        super();
        _CliIssueCommand_instances.add(this);
        __classPrivateFieldGet(this, _CliIssueCommand_instances, "m", _CliIssueCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _CliIssueCommand_instances, "m", _CliIssueCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _CliIssueCommand_instances, "m", _CliIssueCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        let issueLink = '';
        switch (encodeURIComponent(args.options.type)) {
            case 'bug':
                issueLink = 'https://aka.ms/cli-m365/bug';
                break;
            case 'command':
                issueLink = 'https://aka.ms/cli-m365/new-command';
                break;
            case 'sample':
                issueLink = 'https://aka.ms/cli-m365/new-sample-script';
                break;
        }
        await browserUtil.open(issueLink);
        await logger.log(issueLink);
    }
}
_a = CliIssueCommand, _CliIssueCommand_instances = new WeakSet(), _CliIssueCommand_initTelemetry = function _CliIssueCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            type: args.options.type
        });
    });
}, _CliIssueCommand_initOptions = function _CliIssueCommand_initOptions() {
    this.options.unshift({
        option: '-t, --type <type>',
        autocomplete: _a.issueType
    });
}, _CliIssueCommand_initValidators = function _CliIssueCommand_initValidators() {
    this.validators.push(async (args) => {
        if (_a.issueType.indexOf(args.options.type) < 0) {
            return `${args.options.type} is not a valid Issue type. Allowed values are ${_a.issueType.join(', ')}`;
        }
        return true;
    });
};
CliIssueCommand.issueType = [
    'bug',
    'command',
    'sample'
];
export default new CliIssueCommand();
//# sourceMappingURL=cli-issue.js.map