import { JsonRule } from '../../JsonRule.js';
import * as spfxDeps from '../spfx-deps.js';
export class FN021008_PKG_no_duplicate_deps extends JsonRule {
    get id() {
        return 'FN021008';
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
        if (!project.version ||
            !project.packageJson ||
            !project.packageJson.dependencies ||
            !project.packageJson.devDependencies) {
            return;
        }
        const projectDeps = Object.keys(project.packageJson.dependencies);
        const projectDevDeps = Object.keys(project.packageJson.devDependencies);
        const duplicateDeps = projectDeps.filter(dep => projectDevDeps.includes(dep));
        if (duplicateDeps.length === 0) {
            return;
        }
        duplicateDeps.forEach(dep => {
            const isDevDep = this.isDevDep(dep);
            const nodeToUninstall = this.getAstNodeFromFile(project.packageJson, `${isDevDep ? 'dependencies' : 'devDependencies'}.${dep}`);
            this.addFindingWithCustomInfo(`Duplicate ${dep} installed in the project`, `Duplicate ${dep} installed in the project. Install it only as a ${isDevDep ? 'devDependency' : 'dependency'}`, [{
                    file: this.file,
                    resolution: `${isDevDep ? 'installDev' : 'install'} ${dep}@${isDevDep ? project.packageJson.devDependencies[dep] : project.packageJson.dependencies[dep]}`,
                    position: this.getPositionFromNode(nodeToUninstall)
                }], findings);
        });
    }
    isDevDep(dep) {
        if (dep.indexOf('@types/') === 0) {
            return true;
        }
        if (spfxDeps.devDeps.includes(dep)) {
            return true;
        }
        return false;
    }
}
//# sourceMappingURL=FN021008_PKG_no_duplicate_deps.js.map