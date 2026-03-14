var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpfxProjectDoctorCommand_instances, _a, _SpfxProjectDoctorCommand_initTelemetry, _SpfxProjectDoctorCommand_initOptions, _SpfxProjectDoctorCommand_initValidators;
import fs from 'fs';
import os from 'os';
import path from 'path';
import { CommandError } from '../../../../Command.js';
import { packageManager } from '../../../../utils/packageManager.js';
import commands from '../../commands.js';
import { BaseProjectCommand } from './base-project-command.js';
import { rules as genericRules } from './project-doctor/generic-rules.js';
import { FN017001_MISC_npm_dedupe } from './project-upgrade/rules/FN017001_MISC_npm_dedupe.js';
class SpfxProjectDoctorCommand extends BaseProjectCommand {
    get allowedOutputs() {
        return ['json', 'text', 'md', 'tour'];
    }
    get name() {
        return commands.PROJECT_DOCTOR;
    }
    get description() {
        return 'Validates correctness of a SharePoint Framework project';
    }
    constructor() {
        super();
        _SpfxProjectDoctorCommand_instances.add(this);
        this.allFindings = [];
        this.packageManager = 'npm';
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
            '1.9.1',
            '1.10.0',
            '1.11.0',
            '1.12.0',
            '1.12.1',
            '1.13.0',
            '1.13.1',
            '1.14.0',
            '1.15.0',
            '1.15.2',
            '1.16.0',
            '1.16.1',
            '1.17.0',
            '1.17.1',
            '1.17.2',
            '1.17.3',
            '1.17.4',
            '1.18.0',
            '1.18.1',
            '1.18.2',
            '1.19.0',
            '1.20.0',
            '1.21.0',
            '1.21.1',
            '1.22.0',
            '1.22.1',
            '1.22.2'
        ];
        __classPrivateFieldGet(this, _SpfxProjectDoctorCommand_instances, "m", _SpfxProjectDoctorCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpfxProjectDoctorCommand_instances, "m", _SpfxProjectDoctorCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpfxProjectDoctorCommand_instances, "m", _SpfxProjectDoctorCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        this.projectRootPath = this.getProjectRoot(process.cwd());
        if (this.projectRootPath === null) {
            throw new CommandError(`Couldn't find project root folder`, _a.ERROR_NO_PROJECT_ROOT_FOLDER);
        }
        this.packageManager = args.options.packageManager || 'npm';
        if (this.verbose) {
            await logger.logToStderr('Collecting project...');
        }
        const project = this.getProject(this.projectRootPath);
        if (this.debug) {
            await logger.logToStderr('Collected project');
            await logger.logToStderr(project);
        }
        project.version = this.getProjectVersion();
        if (!project.version) {
            throw new CommandError(`Unable to determine the version of the current SharePoint Framework project`, _a.ERROR_NO_VERSION);
        }
        if (!this.supportedVersions.includes(project.version)) {
            throw new CommandError(`CLI for Microsoft 365 doesn't support validating projects built using SharePoint Framework v${project.version}`, _a.ERROR_UNSUPPORTED_VERSION);
        }
        if (this.verbose) {
            await logger.logToStderr(`Project built using SPFx v${project.version}`);
        }
        const rules = [...genericRules];
        try {
            const versionRules = (await import(`./project-doctor/doctor-${project.version}.js`)).default;
            rules.push(...versionRules);
        }
        catch (e) {
            throw new CommandError(e.message);
        }
        rules.forEach(r => {
            r.visit(project, this.allFindings);
        });
        if (this.packageManager === 'npm') {
            const npmDedupeRule = new FN017001_MISC_npm_dedupe();
            npmDedupeRule.visit(project, this.allFindings);
        }
        // remove superseded findings
        this.allFindings
            // get findings that supersede other findings
            .filter(f => f.supersedes.length > 0)
            .forEach(f => {
            f.supersedes.forEach(s => {
                // find the superseded finding
                const i = this.allFindings.findIndex(f1 => f1.id === s);
                if (i > -1) {
                    // ...and remove it from findings
                    this.allFindings.splice(i, 1);
                }
            });
        });
        // flatten
        const findingsToReport = [].concat.apply([], this.allFindings.map(f => {
            return f.occurrences.map(o => {
                return {
                    description: f.description,
                    id: f.id,
                    file: o.file,
                    position: o.position,
                    resolution: o.resolution,
                    resolutionType: f.resolutionType,
                    severity: f.severity,
                    title: f.title
                };
            });
        }));
        // replace package operation tokens with command for the specific package manager
        findingsToReport.forEach(f => {
            // matches must be in this particular order to avoid false matches, eg.
            // uninstallDev contains install
            if (f.resolution.startsWith('uninstallDev')) {
                f.resolution = f.resolution.replace('uninstallDev', packageManager.getPackageManagerCommand('uninstallDev', this.packageManager));
                return;
            }
            if (f.resolution.startsWith('installDev')) {
                f.resolution = f.resolution.replace('installDev', packageManager.getPackageManagerCommand('installDev', this.packageManager));
                return;
            }
            if (f.resolution.startsWith('uninstall')) {
                f.resolution = f.resolution.replace('uninstall', packageManager.getPackageManagerCommand('uninstall', this.packageManager));
                return;
            }
            if (f.resolution.startsWith('install')) {
                f.resolution = f.resolution.replace('install', packageManager.getPackageManagerCommand('install', this.packageManager));
                return;
            }
        });
        switch (args.options.output) {
            case 'text':
                await logger.log(this.getTextReport(findingsToReport));
                break;
            case 'tour':
                this.writeReportTourFolder(this.getTourReport(findingsToReport));
                break;
            case 'md':
                await logger.log(this.getMdReport(findingsToReport));
                break;
            default:
                await logger.log(findingsToReport);
        }
    }
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    getMdOutput(logStatement, command, options) {
        // overwrite markdown output to return the output as-is
        // because the command already implements its own logic to format the output
        return logStatement;
    }
    writeReportTourFolder(findingsToReport) {
        const toursFolder = path.join(this.projectRootPath, '.tours');
        if (!fs.existsSync(toursFolder)) {
            fs.mkdirSync(toursFolder, { recursive: false });
        }
        const tourFilePath = path.join(this.projectRootPath, '.tours', 'validation.tour');
        fs.writeFileSync(path.resolve(tourFilePath), findingsToReport, 'utf-8');
    }
    getTextReport(findings) {
        if (findings.length === 0) {
            return '✅ CLI for Microsoft 365 has found no issues in your project';
        }
        const reportData = this.getReportData(findings);
        const s = [
            'Execute in command line', os.EOL,
            '-----------------------', os.EOL,
            reportData.packageManagerCommands.join(os.EOL), os.EOL,
            os.EOL
        ];
        return s.join('').trim();
    }
    getMdReport(findings) {
        const projectName = this.getProject(this.projectRootPath).packageSolutionJson?.solution?.name;
        const findingsToReport = [];
        const reportData = this.getReportData(findings);
        findings.forEach(f => {
            let resolution = '';
            switch (f.resolutionType) {
                case 'cmd':
                    resolution = `Execute the following command:

\`\`\`sh
${f.resolution}
\`\`\`
`;
                    break;
            }
            findingsToReport.push(`### ${f.id} ${f.title} | ${f.severity}`, os.EOL, os.EOL, f.description, os.EOL, os.EOL, resolution, os.EOL, `File: [${f.file}${(f.position ? `:${f.position.line}:${f.position.character}` : '')}](${f.file})`, os.EOL, os.EOL);
        });
        const s = [
            `# Validate project ${projectName}`, os.EOL,
            os.EOL,
            `Date: ${(new Date().toLocaleDateString())}`, os.EOL,
            os.EOL,
            '## Findings', os.EOL,
            os.EOL
        ];
        if (findingsToReport.length === 0) {
            s.push(`✅ CLI for Microsoft 365 has found no issues in your project`, os.EOL);
        }
        else {
            s.push(...[
                `Following is the list of issues found in your project. [Summary](#Summary) of the recommended fixes is included at the end of the report.`, os.EOL,
                os.EOL,
                findingsToReport.join(''),
                '## Summary', os.EOL,
                os.EOL,
                '### Execute script', os.EOL,
                os.EOL,
                '```sh', os.EOL,
                reportData.packageManagerCommands.join(os.EOL), os.EOL,
                '```', os.EOL,
                os.EOL
            ]);
        }
        return s.join('').trim();
    }
    getTourReport(findings) {
        const projectName = this.getProject(this.projectRootPath).packageSolutionJson?.solution?.name;
        const tourFindings = {
            title: `Validate project ${projectName}`,
            steps: []
        };
        findings.forEach(f => {
            const lineNumber = f.position && f.position.line ? f.position.line : 1;
            let resolution = '';
            switch (f.resolutionType) {
                case 'cmd':
                    resolution = `Execute the following command:\r\n\r\n[\`${f.resolution}\`](command:codetour.sendTextToTerminal?["${f.resolution}"])`;
                    break;
            }
            // Make severity uppercase for the markdown
            const sev = f.severity.toUpperCase();
            // Clean up the file name
            let file = f.file;
            if (file !== undefined) {
                // CodeTour expects the files to be relative from root (i.e.: no './')
                file = file.replace(/\.\//g, '');
                // CodeTour also expects forward slashes as directory separators
                file = file.replace(/\\/g, '/');
            }
            // Create a tour step entry
            const step = {
                file,
                title: `${sev}: ${f.title} (${f.id})`,
                description: `### ${sev}\r\n\r\n${f.description}\r\n\r\n${resolution}`,
                line: lineNumber
            };
            tourFindings.steps.push(step);
        });
        // Add the finale
        tourFindings.steps.push({
            file: ".tours/validation.tour",
            title: "RECOMMENDED: Delete tour",
            description: "### THAT'S IT!!!\r\nOnce you have tested that your project has no more issues, you can delete the `.tour` folder and its contents. Otherwise, you'll be prompted to launch this CodeTour every time you open this project."
        });
        return JSON.stringify(tourFindings, null, 2);
    }
    getReportData(findings) {
        const commandsToExecute = [];
        const modificationPerFile = {};
        const modificationTypePerFile = {};
        const packagesDevExact = [];
        const packagesDepExact = [];
        const packagesDepUn = [];
        const packagesDevUn = [];
        findings.forEach(f => {
            packageManager.mapPackageManagerCommand({
                command: f.resolution,
                packagesDevExact,
                packagesDepExact,
                packagesDepUn,
                packagesDevUn,
                packageMgr: this.packageManager
            });
        });
        const packageManagerCommands = packageManager.reducePackageManagerCommand({
            packagesDepExact,
            packagesDevExact,
            packagesDepUn,
            packagesDevUn,
            packageMgr: this.packageManager
        });
        if (this.packageManager === 'npm') {
            const dedupeFinding = findings.filter(f => f.id === 'FN017001');
            if (dedupeFinding.length > 0) {
                packageManagerCommands.push(dedupeFinding[0].resolution);
            }
        }
        return {
            commandsToExecute: commandsToExecute,
            packageManagerCommands: packageManagerCommands,
            modificationPerFile: modificationPerFile,
            modificationTypePerFile: modificationTypePerFile
        };
    }
}
_a = SpfxProjectDoctorCommand, _SpfxProjectDoctorCommand_instances = new WeakSet(), _SpfxProjectDoctorCommand_initTelemetry = function _SpfxProjectDoctorCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            packageManager: args.options.packageManager || 'npm'
        });
    });
}, _SpfxProjectDoctorCommand_initOptions = function _SpfxProjectDoctorCommand_initOptions() {
    this.options.forEach(o => {
        if (o.option.indexOf('--output') > -1) {
            o.autocomplete = this.allowedOutputs;
        }
    });
    this.options.unshift({
        option: '--packageManager [packageManager]',
        autocomplete: _a.packageManagers
    });
}, _SpfxProjectDoctorCommand_initValidators = function _SpfxProjectDoctorCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.packageManager) {
            if (_a.packageManagers.indexOf(args.options.packageManager) < 0) {
                return `${args.options.packageManager} is not a supported package manager. Supported package managers are ${_a.packageManagers.join(', ')}`;
            }
        }
        return true;
    });
};
SpfxProjectDoctorCommand.packageManagers = ['npm', 'pnpm', 'yarn'];
SpfxProjectDoctorCommand.ERROR_NO_PROJECT_ROOT_FOLDER = 1;
SpfxProjectDoctorCommand.ERROR_NO_VERSION = 3;
SpfxProjectDoctorCommand.ERROR_UNSUPPORTED_VERSION = 4;
export default new SpfxProjectDoctorCommand();
//# sourceMappingURL=project-doctor.js.map