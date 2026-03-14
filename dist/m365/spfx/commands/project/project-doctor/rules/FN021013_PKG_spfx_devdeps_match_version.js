import { satisfies } from 'semver';
import { JsonRule } from '../../JsonRule.js';
import * as spfxDeps from '../spfx-deps.js';
export class FN021013_PKG_spfx_devdeps_match_version extends JsonRule {
    constructor(version) {
        super();
        this.version = version;
    }
    get id() {
        return 'FN021013';
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
        if (!project.packageJson || !project.packageJson.devDependencies) {
            return;
        }
        if (project.packageJson.devDependencies) {
            const projectDevDeps = Object.keys(project.packageJson.devDependencies);
            this.validateDependencies({
                dependencies: projectDevDeps,
                spfxDeps: spfxDeps.devDeps,
                project,
                version: this.version,
                findings
            });
        }
    }
    validateDependencies({ dependencies, spfxDeps, project, version, findings }) {
        dependencies.forEach(dep => {
            const depVersion = project.packageJson.devDependencies[dep];
            if (!spfxDeps.includes(dep) ||
                satisfies(version, depVersion)) {
                return;
            }
            const node = this.getAstNodeFromFile(project.packageJson, 'devDependencies');
            this.addFindingWithCustomInfo(`${dep} doesn't match project version`, `${dep}@${depVersion} doesn't match the project version ${project.version}`, [{
                    file: this.file,
                    resolution: `installDev ${dep}@${project.version}`,
                    position: this.getPositionFromNode(node)
                }], findings);
        });
    }
}
//# sourceMappingURL=FN021013_PKG_spfx_devdeps_match_version.js.map