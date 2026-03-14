var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpfxProjectGithubWorkflowAddCommand_instances, _a, _SpfxProjectGithubWorkflowAddCommand_initTelemetry, _SpfxProjectGithubWorkflowAddCommand_initOptions, _SpfxProjectGithubWorkflowAddCommand_initValidators;
import fs from 'fs';
import path from 'path';
import yaml from 'yaml';
import { CommandError } from '../../../../Command.js';
import { fsUtil } from '../../../../utils/fsUtil.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import { workflow } from './DeployWorkflow.js';
import { BaseProjectCommand } from './base-project-command.js';
class SpfxProjectGithubWorkflowAddCommand extends BaseProjectCommand {
    get name() {
        return commands.PROJECT_GITHUB_WORKFLOW_ADD;
    }
    get description() {
        return 'Adds a GitHub workflow for a SharePoint Framework project.';
    }
    constructor() {
        super();
        _SpfxProjectGithubWorkflowAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpfxProjectGithubWorkflowAddCommand_instances, "m", _SpfxProjectGithubWorkflowAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpfxProjectGithubWorkflowAddCommand_instances, "m", _SpfxProjectGithubWorkflowAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpfxProjectGithubWorkflowAddCommand_instances, "m", _SpfxProjectGithubWorkflowAddCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        this.projectRootPath = this.getProjectRoot(process.cwd());
        if (this.projectRootPath === null) {
            throw new CommandError(`Couldn't find project root folder`, _a.ERROR_NO_PROJECT_ROOT_FOLDER);
        }
        const solutionPackageJsonFile = path.join(this.projectRootPath, 'package.json');
        const packageJson = fs.readFileSync(solutionPackageJsonFile, 'utf-8');
        const solutionName = JSON.parse(packageJson).name;
        if (this.debug) {
            await logger.logToStderr(`Adding GitHub workflow in the current SPFx project`);
        }
        try {
            this.updateWorkflow(solutionName, workflow, args.options);
            this.saveWorkflow(workflow);
        }
        catch (error) {
            throw new CommandError(error);
        }
    }
    saveWorkflow(workflow) {
        const githubPath = path.join(this.projectRootPath, '.github');
        fsUtil.ensureDirectory(githubPath);
        const workflowPath = path.join(githubPath, 'workflows');
        fsUtil.ensureDirectory(workflowPath);
        const workflowFile = path.join(workflowPath, 'deploy-spfx-solution.yml');
        fs.writeFileSync(path.resolve(workflowFile), yaml.stringify(workflow), 'utf-8');
    }
    updateWorkflow(solutionName, workflow, options) {
        workflow.name = options.name ? options.name : workflow.name.replace('{{ name }}', solutionName);
        if (options.branchName) {
            workflow.on.push.branches[0] = options.branchName;
        }
        if (options.manuallyTrigger) {
            // eslint-disable-next-line camelcase
            workflow.on.workflow_dispatch = null;
        }
        if (options.skipFeatureDeployment) {
            this.getDeployAction(workflow).with.SKIP_FEATURE_DEPLOYMENT = true;
        }
        if (options.loginMethod === 'user') {
            const loginAction = this.getLoginAction(workflow);
            loginAction.with = {
                ADMIN_USERNAME: '${{ secrets.ADMIN_USERNAME }}',
                ADMIN_PASSWORD: '${{ secrets.ADMIN_PASSWORD }}'
            };
        }
        if (options.scope === 'sitecollection') {
            const deployAction = this.getDeployAction(workflow);
            deployAction.with.SCOPE = 'sitecollection';
            deployAction.with.SITE_COLLECTION_URL = options.siteUrl;
        }
        if (solutionName) {
            const deployAction = this.getDeployAction(workflow);
            deployAction.with.APP_FILE_PATH = deployAction.with.APP_FILE_PATH.replace('{{ solutionName }}', solutionName);
        }
    }
    getLoginAction(workflow) {
        const steps = this.getWorkFlowSteps(workflow);
        return steps.find(step => step.uses && step.uses.indexOf('action-cli-login') >= 0);
    }
    getDeployAction(workflow) {
        const steps = this.getWorkFlowSteps(workflow);
        return steps.find(step => step.uses && step.uses.indexOf('action-cli-deploy') >= 0);
    }
    getWorkFlowSteps(workflow) {
        return workflow.jobs['build-and-deploy'].steps;
    }
}
_a = SpfxProjectGithubWorkflowAddCommand, _SpfxProjectGithubWorkflowAddCommand_instances = new WeakSet(), _SpfxProjectGithubWorkflowAddCommand_initTelemetry = function _SpfxProjectGithubWorkflowAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            name: typeof args.options.name !== 'undefined',
            branchName: typeof args.options.branchName !== 'undefined',
            manuallyTrigger: !!args.options.manuallyTrigger,
            loginMethod: typeof args.options.loginMethod !== 'undefined',
            scope: typeof args.options.scope !== 'undefined',
            skipFeatureDeployment: !!args.options.skipFeatureDeployment
        });
    });
}, _SpfxProjectGithubWorkflowAddCommand_initOptions = function _SpfxProjectGithubWorkflowAddCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name [name]'
    }, {
        option: '-b, --branchName [branchName]'
    }, {
        option: '-m, --manuallyTrigger'
    }, {
        option: '-l, --loginMethod [loginMethod]',
        autocomplete: _a.loginMethod
    }, {
        option: '-s, --scope [scope]',
        autocomplete: _a.scope
    }, {
        option: '-u, --siteUrl [siteUrl]'
    }, {
        option: '--skipFeatureDeployment'
    });
}, _SpfxProjectGithubWorkflowAddCommand_initValidators = function _SpfxProjectGithubWorkflowAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.scope && args.options.scope === 'sitecollection') {
            if (!args.options.siteUrl) {
                return `siteUrl option has to be defined when scope set to ${args.options.scope}`;
            }
            const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.siteUrl);
            if (isValidSharePointUrl !== true) {
                return isValidSharePointUrl;
            }
        }
        if (args.options.loginMethod && _a.loginMethod.indexOf(args.options.loginMethod) < 0) {
            return `${args.options.loginMethod} is not a valid login method. Allowed values are ${_a.loginMethod.join(', ')}`;
        }
        if (args.options.scope && _a.scope.indexOf(args.options.scope) < 0) {
            return `${args.options.scope} is not a valid scope. Allowed values are ${_a.scope.join(', ')}`;
        }
        return true;
    });
};
SpfxProjectGithubWorkflowAddCommand.loginMethod = ['application', 'user'];
SpfxProjectGithubWorkflowAddCommand.scope = ['tenant', 'sitecollection'];
SpfxProjectGithubWorkflowAddCommand.ERROR_NO_PROJECT_ROOT_FOLDER = 1;
export default new SpfxProjectGithubWorkflowAddCommand();
//# sourceMappingURL=project-github-workflow-add.js.map