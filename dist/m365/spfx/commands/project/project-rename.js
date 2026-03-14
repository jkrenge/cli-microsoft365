var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpfxProjectRenameCommand_instances, _SpfxProjectRenameCommand_initTelemetry, _SpfxProjectRenameCommand_initOptions;
import fs from 'fs';
import path from 'path';
import { v4 } from 'uuid';
import { CommandError } from '../../../../Command.js';
import commands from '../../commands.js';
import { BaseProjectCommand } from './base-project-command.js';
class SpfxProjectRenameCommand extends BaseProjectCommand {
    get name() {
        return commands.PROJECT_RENAME;
    }
    get description() {
        return 'Renames SharePoint Framework project';
    }
    constructor() {
        super();
        _SpfxProjectRenameCommand_instances.add(this);
        this.generateNewId = () => {
            return v4();
        };
        this.replacePackageJsonContent = async (filePath, args, logger) => {
            if (!fs.existsSync(filePath)) {
                return;
            }
            const existingContent = fs.readFileSync(filePath, 'utf-8');
            const updatedContent = JSON.parse(existingContent);
            if (updatedContent &&
                updatedContent.name) {
                updatedContent.name = args.options.newName;
            }
            const updatedContentString = JSON.stringify(updatedContent, null, 2);
            if (updatedContentString !== existingContent) {
                fs.writeFileSync(filePath, updatedContentString, 'utf-8');
                if (this.debug) {
                    await logger.logToStderr(`Updated ${path.basename(filePath)}`);
                }
            }
        };
        this.replaceYoRcJsonContent = async (filePath, newId, args, logger) => {
            if (!fs.existsSync(filePath)) {
                return;
            }
            const existingContent = fs.readFileSync(filePath, 'utf-8');
            const updatedContent = JSON.parse(existingContent);
            if (updatedContent &&
                updatedContent['@microsoft/generator-sharepoint'] &&
                updatedContent['@microsoft/generator-sharepoint'].libraryName) {
                updatedContent['@microsoft/generator-sharepoint'].libraryName = args.options.newName;
            }
            if (updatedContent &&
                updatedContent['@microsoft/generator-sharepoint'] &&
                updatedContent['@microsoft/generator-sharepoint'].solutionName) {
                updatedContent['@microsoft/generator-sharepoint'].solutionName = args.options.newName;
            }
            if (updatedContent &&
                updatedContent['@microsoft/generator-sharepoint'] &&
                updatedContent['@microsoft/generator-sharepoint'].libraryId &&
                args.options.generateNewId) {
                updatedContent['@microsoft/generator-sharepoint'].libraryId = newId;
            }
            const updatedContentString = JSON.stringify(updatedContent, null, 2);
            if (updatedContentString !== existingContent) {
                fs.writeFileSync(filePath, updatedContentString, 'utf-8');
                if (this.debug) {
                    await logger.logToStderr(`Updated ${path.basename(filePath)}`);
                }
            }
        };
        this.replacePackageSolutionJsonContent = async (filePath, projectName, newId, args, logger) => {
            if (!fs.existsSync(filePath)) {
                return;
            }
            const existingContent = fs.readFileSync(filePath, 'utf-8');
            const updatedContent = JSON.parse(existingContent);
            if (updatedContent &&
                updatedContent.solution &&
                updatedContent.solution.name) {
                updatedContent.solution.name = updatedContent.solution.name.replace(new RegExp(projectName, 'g'), args.options.newName);
            }
            if (updatedContent &&
                updatedContent.solution &&
                updatedContent.solution.id &&
                args.options.generateNewId) {
                updatedContent.solution.id = newId;
            }
            if (updatedContent &&
                updatedContent.paths &&
                updatedContent.paths.zippedPackage) {
                updatedContent.paths.zippedPackage = updatedContent.paths.zippedPackage.replace(new RegExp(projectName, 'g'), args.options.newName);
            }
            const updatedContentString = JSON.stringify(updatedContent, null, 2);
            if (updatedContentString !== existingContent) {
                fs.writeFileSync(filePath, updatedContentString, 'utf-8');
                if (this.debug) {
                    await logger.logToStderr(`Updated ${path.basename(filePath)}`);
                }
            }
        };
        this.replaceDeployAzureStorageJsonContent = async (filePath, args, logger) => {
            if (!fs.existsSync(filePath)) {
                return;
            }
            const existingContent = fs.readFileSync(filePath, 'utf-8');
            const updatedContent = JSON.parse(existingContent);
            if (updatedContent &&
                updatedContent.container) {
                updatedContent.container = args.options.newName;
            }
            const updatedContentString = JSON.stringify(updatedContent, null, 2);
            if (updatedContentString !== existingContent) {
                fs.writeFileSync(filePath, updatedContentString, 'utf-8');
                if (this.debug) {
                    await logger.logToStderr(`Updated ${path.basename(filePath)}`);
                }
            }
        };
        this.replaceReadMeContent = async (filePath, projectName, args, logger) => {
            if (!fs.existsSync(filePath)) {
                return;
            }
            const existingContent = fs.readFileSync(filePath, 'utf-8');
            const updatedContent = existingContent.replace(new RegExp(projectName, 'g'), args.options.newName);
            if (updatedContent !== existingContent) {
                fs.writeFileSync(filePath, updatedContent, 'utf-8');
                if (this.debug) {
                    await logger.logToStderr(`Updated ${path.basename(filePath)}`);
                }
            }
        };
        __classPrivateFieldGet(this, _SpfxProjectRenameCommand_instances, "m", _SpfxProjectRenameCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpfxProjectRenameCommand_instances, "m", _SpfxProjectRenameCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        this.projectRootPath = this.getProjectRoot(process.cwd());
        if (this.projectRootPath === null) {
            throw new CommandError(`Couldn't find project root folder`, SpfxProjectRenameCommand.ERROR_NO_PROJECT_ROOT_FOLDER);
        }
        const packageJson = this.getProject(this.projectRootPath).packageJson;
        const projectName = packageJson.name;
        let newId = '';
        if (args.options.generateNewId) {
            newId = this.generateNewId();
            if (this.debug) {
                await logger.logToStderr('Created new solution id');
                await logger.logToStderr(newId);
            }
        }
        if (this.debug) {
            await logger.logToStderr(`Renaming SharePoint Framework project to '${args.options.newName}'`);
        }
        try {
            await this.replacePackageJsonContent(path.join(this.projectRootPath, 'package.json'), args, logger);
            await this.replaceYoRcJsonContent(path.join(this.projectRootPath, '.yo-rc.json'), newId, args, logger);
            await this.replacePackageSolutionJsonContent(path.join(this.projectRootPath, 'config', 'package-solution.json'), projectName, newId, args, logger);
            await this.replaceDeployAzureStorageJsonContent(path.join(this.projectRootPath, 'config', 'deploy-azure-storage.json'), args, logger);
            await this.replaceReadMeContent(path.join(this.projectRootPath, 'README.md'), projectName, args, logger);
        }
        catch (error) {
            throw new CommandError(error);
        }
    }
}
_SpfxProjectRenameCommand_instances = new WeakSet(), _SpfxProjectRenameCommand_initTelemetry = function _SpfxProjectRenameCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            generateNewId: args.options.generateNewId
        });
    });
}, _SpfxProjectRenameCommand_initOptions = function _SpfxProjectRenameCommand_initOptions() {
    this.options.unshift({
        option: '-n, --newName <newName>'
    }, {
        option: '--generateNewId'
    });
};
SpfxProjectRenameCommand.ERROR_NO_PROJECT_ROOT_FOLDER = 1;
export default new SpfxProjectRenameCommand();
//# sourceMappingURL=project-rename.js.map