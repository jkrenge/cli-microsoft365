var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpfxProjectAzureDevOpsPipelineAddCommand_instances, _a, _SpfxProjectAzureDevOpsPipelineAddCommand_initTelemetry, _SpfxProjectAzureDevOpsPipelineAddCommand_initOptions, _SpfxProjectAzureDevOpsPipelineAddCommand_initValidators;
import fs from 'fs';
import path from 'path';
import yaml from 'yaml';
import { CommandError } from '../../../../Command.js';
import commands from '../../commands.js';
import { BaseProjectCommand } from './base-project-command.js';
import { validation } from '../../../../utils/validation.js';
import { pipeline } from './DeployWorkflow.js';
import { fsUtil } from '../../../../utils/fsUtil.js';
class SpfxProjectAzureDevOpsPipelineAddCommand extends BaseProjectCommand {
    get name() {
        return commands.PROJECT_AZUREDEVOPS_PIPELINE_ADD;
    }
    get description() {
        return 'Adds a Azure DevOps Pipeline for a SharePoint Framework project.';
    }
    constructor() {
        super();
        _SpfxProjectAzureDevOpsPipelineAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpfxProjectAzureDevOpsPipelineAddCommand_instances, "m", _SpfxProjectAzureDevOpsPipelineAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpfxProjectAzureDevOpsPipelineAddCommand_instances, "m", _SpfxProjectAzureDevOpsPipelineAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpfxProjectAzureDevOpsPipelineAddCommand_instances, "m", _SpfxProjectAzureDevOpsPipelineAddCommand_initValidators).call(this);
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
            await logger.logToStderr(`Adding Azure DevOps pipeline in the current SPFx project`);
        }
        try {
            this.updatePipeline(solutionName, pipeline, args.options);
            this.savePipeline(pipeline);
        }
        catch (error) {
            throw new CommandError(error);
        }
    }
    savePipeline(pipeline) {
        const azureDevOpsPath = path.join(this.projectRootPath, '.azuredevops');
        fsUtil.ensureDirectory(azureDevOpsPath);
        const pipelinesPath = path.join(azureDevOpsPath, 'pipelines');
        fsUtil.ensureDirectory(pipelinesPath);
        const pipelineFile = path.join(pipelinesPath, 'deploy-spfx-solution.yml');
        fs.writeFileSync(path.resolve(pipelineFile), yaml.stringify(pipeline), 'utf-8');
    }
    updatePipeline(solutionName, pipeline, options) {
        if (options.name) {
            pipeline.name = options.name;
        }
        else {
            delete pipeline.name;
        }
        if (options.branchName) {
            pipeline.trigger.branches.include[0] = options.branchName;
        }
        const script = this.getScriptAction(pipeline);
        if (script.script) {
            if (options.loginMethod === 'user') {
                script.script = script.script.replace(`{{login}}`, `m365 login --authType password --userName '$(UserName)' --password '$(Password)'`);
                pipeline.variables = pipeline.variables.filter(v => v.name !== 'CertificateBase64Encoded' &&
                    v.name !== 'CertificateSecureFileId' &&
                    v.name !== 'CertificatePassword' &&
                    v.name !== 'EntraAppId' &&
                    v.name !== 'TenantId');
            }
            else {
                script.script = script.script.replace(`{{login}}`, `m365 login --authType certificate --certificateBase64Encoded '$(CertificateBase64Encoded)' --password '$(CertificatePassword)' --appId '$(EntraAppId)' --tenant '$(TenantId)'`);
                pipeline.variables = pipeline.variables.filter(v => v.name !== 'UserName' &&
                    v.name !== 'Password');
            }
            if (options.scope === 'sitecollection') {
                script.script = script.script.replace(`{{deploy}}`, `m365 spo app deploy --name '$(PackageName)' --appCatalogScope sitecollection --appCatalogUrl '$(SiteAppCatalogUrl)'`);
                script.script = script.script.replace(`{{addApp}}`, `m365 spo app add --filePath '$(Build.SourcesDirectory)/sharepoint/solution/$(PackageName)' --appCatalogScope sitecollection --appCatalogUrl '$(SiteAppCatalogUrl)' --overwrite`);
                this.assignPipelineVariables(pipeline, 'SiteAppCatalogUrl', options.siteUrl);
            }
            else {
                script.script = script.script.replace(`{{deploy}}`, `m365 spo app deploy --name '$(PackageName)' --appCatalogScope 'tenant'`);
                script.script = script.script.replace(`{{addApp}}`, `m365 spo app add --filePath '$(Build.SourcesDirectory)/sharepoint/solution/$(PackageName)' --overwrite`);
                pipeline.variables = pipeline.variables.filter(v => v.name !== 'SiteAppCatalogUrl');
            }
            if (solutionName) {
                this.assignPipelineVariables(pipeline, 'PackageName', `${solutionName}.sppkg`);
            }
            if (options.skipFeatureDeployment) {
                script.script = script.script.replace(`m365 spo app deploy `, `m365 spo app deploy --skipFeatureDeployment `);
            }
        }
    }
    assignPipelineVariables(pipeline, variableName, newVariableValue) {
        const variable = pipeline.variables.find(v => v.name === variableName);
        if (variable) {
            variable.value = newVariableValue;
        }
    }
    getScriptAction(pipeline) {
        const steps = this.getPipelineSteps(pipeline);
        return steps.find(step => step.script);
    }
    getPipelineSteps(pipeline) {
        return pipeline.stages[0].jobs[0].steps;
    }
}
_a = SpfxProjectAzureDevOpsPipelineAddCommand, _SpfxProjectAzureDevOpsPipelineAddCommand_instances = new WeakSet(), _SpfxProjectAzureDevOpsPipelineAddCommand_initTelemetry = function _SpfxProjectAzureDevOpsPipelineAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            name: typeof args.options.name !== 'undefined',
            branchName: typeof args.options.branchName !== 'undefined',
            loginMethod: typeof args.options.loginMethod !== 'undefined',
            scope: typeof args.options.scope !== 'undefined',
            skipFeatureDeployment: !!args.options.skipFeatureDeployment
        });
    });
}, _SpfxProjectAzureDevOpsPipelineAddCommand_initOptions = function _SpfxProjectAzureDevOpsPipelineAddCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name [name]'
    }, {
        option: '-b, --branchName [branchName]'
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
}, _SpfxProjectAzureDevOpsPipelineAddCommand_initValidators = function _SpfxProjectAzureDevOpsPipelineAddCommand_initValidators() {
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
SpfxProjectAzureDevOpsPipelineAddCommand.loginMethod = ['application', 'user'];
SpfxProjectAzureDevOpsPipelineAddCommand.scope = ['tenant', 'sitecollection'];
SpfxProjectAzureDevOpsPipelineAddCommand.ERROR_NO_PROJECT_ROOT_FOLDER = 1;
export default new SpfxProjectAzureDevOpsPipelineAddCommand();
//# sourceMappingURL=project-azuredevops-pipeline-add.js.map