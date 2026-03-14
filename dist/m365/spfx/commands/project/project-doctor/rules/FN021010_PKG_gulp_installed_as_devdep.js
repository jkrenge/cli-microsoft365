import { JsonRule } from '../../JsonRule.js';
export class FN021010_PKG_gulp_installed_as_devdep extends JsonRule {
    get id() {
        return 'FN021010';
    }
    get title() {
        return 'gulp installed as a devDependency';
    }
    get description() {
        return 'gulp is installed as a dependency. Install it as a devDependency instead';
    }
    get severity() {
        return 'Required';
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
    visit(project, findings) {
        if (!project.packageJson ||
            !project.packageJson.dependencies) {
            return;
        }
        const projectDeps = Object.keys(project.packageJson.dependencies);
        if (!projectDeps.includes('gulp')) {
            return;
        }
        const node = this.getAstNodeFromFile(project.packageJson, `dependencies.gulp`);
        this.addFindingWithCustomInfo(this.title, this.description, [{
                file: this.file,
                resolution: `installDev gulp@${project.packageJson.dependencies['gulp']}`,
                position: this.getPositionFromNode(node)
            }], findings);
    }
}
//# sourceMappingURL=FN021010_PKG_gulp_installed_as_devdep.js.map