import { satisfies } from 'semver';
import { JsonRule } from '../../JsonRule.js';
import * as spfxDeps from '../spfx-deps.js';
export class FN021001_PKG_spfx_deps_versions_match_project_version extends JsonRule {
    constructor(includeDevDeps = true) {
        super();
        this.includeDevDeps = includeDevDeps;
    }
    get id() {
        return 'FN021001';
    }
    get title() {
        return '';
    }
    get description() {
        return '';
    }
    get severity() {
        return 'Required';
    }
    get file() {
        return './package.json';
    }
    get resolutionType() {
        return 'cmd';
    }
    visit(project, findings) {
        if (!project.version || !project.packageJson) {
            return;
        }
        const allSpfxDeps = spfxDeps.deps.concat(spfxDeps.devDeps);
        if (project.packageJson.dependencies) {
            const projectDeps = Object.keys(project.packageJson.dependencies);
            this.validateDependencies({
                dependencies: projectDeps,
                isDevDep: false,
                allSpfxDeps,
                project,
                findings
            });
        }
        if (this.includeDevDeps && project.packageJson.devDependencies) {
            const projectDevDeps = Object.keys(project.packageJson.devDependencies);
            this.validateDependencies({
                dependencies: projectDevDeps,
                isDevDep: true,
                allSpfxDeps,
                project,
                findings
            });
        }
    }
    validateDependencies({ dependencies, isDevDep, allSpfxDeps, project, findings }) {
        dependencies.forEach(dep => {
            const depVersion = isDevDep ?
                project.packageJson.devDependencies[dep] :
                project.packageJson.dependencies[dep];
            if (!allSpfxDeps.includes(dep) ||
                satisfies(project.version, depVersion)) {
                return;
            }
            const node = this.getAstNodeFromFile(project.packageJson, `${isDevDep ? 'devDependencies' : 'dependencies'}.${dep}`);
            this.addFindingWithCustomInfo(`${dep} doesn't match project version`, `${dep}@${depVersion} doesn't match the project version ${project.version}`, [{
                    file: this.file,
                    resolution: `${isDevDep ? 'installDev' : 'install'} ${dep}@${project.version}`,
                    position: this.getPositionFromNode(node)
                }], findings);
        });
    }
}
//# sourceMappingURL=FN021001_PKG_spfx_deps_versions_match_project_version.js.map