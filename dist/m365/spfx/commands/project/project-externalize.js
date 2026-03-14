var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpfxProjectExternalizeCommand_instances, _SpfxProjectExternalizeCommand_initOptions;
import os from 'os';
import path from 'path';
import { CommandError } from '../../../../Command.js';
import commands from '../../commands.js';
import { BaseProjectCommand } from './base-project-command.js';
import rules from './project-externalize/DefaultRules.js';
class SpfxProjectExternalizeCommand extends BaseProjectCommand {
    get allowedOutputs() {
        return ['json', 'text', 'md'];
    }
    get name() {
        return commands.PROJECT_EXTERNALIZE;
    }
    get description() {
        return 'Externalizes SharePoint Framework project dependencies';
    }
    constructor() {
        super();
        _SpfxProjectExternalizeCommand_instances.add(this);
        this.supportedVersions = [
            '1.0.0',
            '1.0.1',
            '1.0.2',
            '1.1.0',
            '1.1.1',
            '1.1.3',
            '1.2.0',
            '1.3.0',
            '1.3.1',
            '1.3.2',
            '1.3.4',
            '1.4.0',
            '1.4.1',
            '1.5.0',
            '1.5.1',
            '1.6.0',
            '1.7.0',
            '1.7.1',
            '1.8.0',
            '1.8.1',
            '1.8.2',
            '1.9.1'
        ];
        this.allFindings = [];
        this.allEditSuggestions = [];
        __classPrivateFieldGet(this, _SpfxProjectExternalizeCommand_instances, "m", _SpfxProjectExternalizeCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.output !== 'json' || this.verbose) {
            await logger.logToStderr(`This command is currently in preview. Feedback welcome at https://github.com/pnp/cli-microsoft365/issues${os.EOL}`);
        }
        this.projectRootPath = this.getProjectRoot(process.cwd());
        if (this.projectRootPath === null) {
            throw new CommandError(`Couldn't find project root folder`, SpfxProjectExternalizeCommand.ERROR_NO_PROJECT_ROOT_FOLDER);
        }
        this.projectVersion = this.getProjectVersion();
        if (!this.projectVersion) {
            throw new CommandError(`Unable to determine the version of the current SharePoint Framework project`, SpfxProjectExternalizeCommand.ERROR_NO_VERSION);
        }
        if (this.projectVersion && this.supportedVersions.indexOf(this.projectVersion) < 0) {
            throw new CommandError(`CLI for Microsoft 365 doesn't support externalizing dependencies of SharePoint Framework projects of version ${this.projectVersion}. Supported versions are ${this.supportedVersions.join(', ')}`, SpfxProjectExternalizeCommand.ERROR_UNSUPPORTED_VERSION);
        }
        if (this.verbose) {
            await logger.logToStderr('Collecting project...');
        }
        const project = this.getProject(this.projectRootPath);
        if (this.debug) {
            await logger.logToStderr('Collected project');
            await logger.logToStderr(project);
        }
        const asyncRulesResults = rules.map(r => r.visit(project));
        try {
            const rulesResults = await Promise.all(asyncRulesResults);
            this.allFindings.push(...rulesResults.map(x => x.entries).reduce((x, y) => [...x, ...y]));
            this.allEditSuggestions.push(...rulesResults.map(x => x.suggestions).reduce((x, y) => [...x, ...y]));
            //removing duplicates
            this.allFindings = this.allFindings.filter((x, i) => this.allFindings.findIndex(y => y.key === x.key) === i);
            await this.writeReport(this.allFindings, this.allEditSuggestions, logger, args.options);
        }
        catch (err) {
            throw new CommandError(err);
        }
    }
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    getMdOutput(logStatement, command, options) {
        // overwrite markdown output to return the output as-is
        // because the command already implements its own logic to format the output
        return logStatement;
    }
    async writeReport(findingsToReport, editsToReport, logger, options) {
        let report;
        switch (options.output) {
            case 'json':
                report = { externalConfiguration: this.serializeJsonReport(findingsToReport), edits: editsToReport };
                break;
            case 'md':
                report = this.serializeMdReport(findingsToReport, editsToReport);
                break;
            default:
                report = this.serializeTextReport(findingsToReport, editsToReport);
                break;
        }
        await logger.log(report);
    }
    serializeMdReport(findingsToReport, editsToReport) {
        const lines = [
            `# Externalizing dependencies of project ${path.basename(this.projectRootPath)}`, os.EOL,
            os.EOL,
            `Date: ${(new Date().toLocaleDateString())}`, os.EOL,
            os.EOL,
            '## Findings', os.EOL,
            os.EOL,
            '### Modify files', os.EOL,
            os.EOL,
            '#### [config.json](config/config.json)', os.EOL,
            os.EOL,
            'Replace the externals property (or add if not defined) with', os.EOL,
            os.EOL,
            '```json', os.EOL,
            JSON.stringify(this.serializeJsonReport(findingsToReport), null, 2), os.EOL,
            '```', os.EOL,
            ...this.getReportForFileEdit(this.getGroupedFileEdits(editsToReport, 'add')),
            ...this.getReportForFileEdit(this.getGroupedFileEdits(editsToReport, 'remove'))
        ];
        return lines.join('');
    }
    getReportForFileEdit(suggestions) {
        const initialReport = suggestions
            .map(x => [
            `#### [${x[0].path}](${x[0].path})`, os.EOL,
            x[0].action, os.EOL,
            '```JavaScript', os.EOL,
            ...x.map(y => [y.targetValue, os.EOL]).reduce((y, z) => [...y, ...z]), '```', os.EOL
        ]);
        if (initialReport.length > 0) {
            return initialReport.reduce((x, y) => [...x, ...y]);
        }
        else {
            return [];
        }
    }
    getGroupedFileEdits(editsToReport, action) {
        const editsMatchingAction = editsToReport.filter(x => x.action === action);
        return editsMatchingAction
            .filter((x, i) => editsMatchingAction.findIndex(y => y.path === x.path) === i)
            .map(x => editsMatchingAction.filter(y => y.path === x.path));
    }
    serializeJsonReport(findingsToReport) {
        const result = {};
        findingsToReport.forEach((f) => {
            if (!f.globalName) {
                result[f.key] = f.path;
            }
            else {
                result[f.key] = {
                    path: f.path,
                    globalName: f.globalName,
                    globalDependencies: f.globalDependencies
                };
            }
        });
        return {
            externals: result
        };
    }
    serializeTextReport(findingsToReport, editsToReport) {
        const s = [
            'In the config/config.json file update the externals property to:', os.EOL,
            os.EOL,
            JSON.stringify({ externalConfiguration: this.serializeJsonReport(findingsToReport), edits: editsToReport }, null, 2)
        ];
        return s.join('').trim();
    }
}
_SpfxProjectExternalizeCommand_instances = new WeakSet(), _SpfxProjectExternalizeCommand_initOptions = function _SpfxProjectExternalizeCommand_initOptions() {
    this.options.forEach(o => {
        if (o.option.indexOf('--output') > -1) {
            o.autocomplete = this.allowedOutputs;
        }
    });
};
SpfxProjectExternalizeCommand.ERROR_NO_PROJECT_ROOT_FOLDER = 1;
SpfxProjectExternalizeCommand.ERROR_NO_VERSION = 3;
SpfxProjectExternalizeCommand.ERROR_UNSUPPORTED_VERSION = 2;
export default new SpfxProjectExternalizeCommand();
//# sourceMappingURL=project-externalize.js.map