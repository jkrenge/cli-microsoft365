import { coerce, SemVer } from 'semver';
import { JsonRule } from '../../JsonRule.js';
export class FN021009_PKG_no_duplicate_oui_deps extends JsonRule {
    get id() {
        return 'FN021009';
    }
    get title() {
        return '@fluentui/react and office-ui-fabric-react installed in the project';
    }
    get description() {
        return '@fluentui/react and office-ui-fabric-react installed in the project. Consider uninstalling @fluentui/react';
    }
    get severity() {
        return 'Optional';
    }
    get file() {
        return './package.json';
    }
    get resolution() {
        return '';
    }
    get resolutionType() {
        return 'cmd';
    }
    customCondition(project) {
        const ouifSemVer = coerce(project.packageJson?.dependencies?.['office-ui-fabric-react']);
        if (!ouifSemVer) {
            return false;
        }
        // ouif and @fluentui/react are both required starting from 7.199.x
        return ouifSemVer < new SemVer('7.199.0');
    }
    visit(project, findings) {
        if (!project.packageJson || !this.customCondition(project)) {
            return;
        }
        const projectDeps = [];
        const projectDevDeps = [];
        if (project.packageJson.dependencies) {
            projectDeps.push(...Object.keys(project.packageJson.dependencies));
        }
        if (project.packageJson.devDependencies) {
            projectDevDeps.push(...Object.keys(project.packageJson.devDependencies));
        }
        const allDeps = projectDeps.concat(projectDevDeps);
        if (!allDeps.includes('office-ui-fabric-react') ||
            !allDeps.includes('@fluentui/react')) {
            return;
        }
        const dependencyType = projectDeps.includes('@fluentui/react') ? 'dependencies' : 'devDependencies';
        const node = this.getAstNodeFromFile(project.packageJson, `${dependencyType}.@fluentui/react`);
        this.addFindingWithCustomInfo(this.title, this.description, [{
                file: this.file,
                resolution: `${dependencyType === 'dependencies' ? 'uninstall' : 'uninstallDev'} @fluentui/react`,
                position: this.getPositionFromNode(node)
            }], findings);
    }
}
//# sourceMappingURL=FN021009_PKG_no_duplicate_oui_deps.js.map