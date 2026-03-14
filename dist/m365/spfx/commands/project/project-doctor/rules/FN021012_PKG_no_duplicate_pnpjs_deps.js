import { JsonRule } from '../../JsonRule.js';
export class FN021012_PKG_no_duplicate_pnpjs_deps extends JsonRule {
    get id() {
        return 'FN021012';
    }
    get title() {
        return 'sp-pnp-js and @pnp/js installed in the project';
    }
    get description() {
        return 'sp-pnp-js and @pnp/js installed in the project. Consider uninstalling the deprecated sp-pnp-js package';
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
    visit(project, findings) {
        if (!project.packageJson) {
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
        if (allDeps.length === 0) {
            return;
        }
        if (!allDeps.includes('sp-pnp-js') ||
            !allDeps.includes('@pnp/sp')) {
            return;
        }
        const dependencyType = projectDeps.includes('sp-pnp-js') ? 'dependencies' : 'devDependencies';
        const node = this.getAstNodeFromFile(project.packageJson, `${dependencyType}.sp-pnp-js`);
        this.addFindingWithCustomInfo(this.title, this.description, [{
                file: this.file,
                resolution: `${dependencyType === 'dependencies' ? 'uninstall' : 'uninstallDev'} sp-pnp-js`,
                position: this.getPositionFromNode(node)
            }], findings);
    }
}
//# sourceMappingURL=FN021012_PKG_no_duplicate_pnpjs_deps.js.map