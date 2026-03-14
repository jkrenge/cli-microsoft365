import { JsonRule } from '../../JsonRule.js';
export class FN021005_PKG_types_installed_as_devdep extends JsonRule {
    get id() {
        return 'FN021005';
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
            !project.packageJson.dependencies) {
            return;
        }
        const projectDeps = Object.keys(project.packageJson.dependencies);
        projectDeps.forEach(dep => {
            if (dep.indexOf('@types/') < 0) {
                return;
            }
            const node = this.getAstNodeFromFile(project.packageJson, `dependencies.${dep}`);
            this.addFindingWithCustomInfo(`${dep} installed as a dependency`, `Package ${dep} is installed as a dependency. Install it as a devDependency instead`, [{
                    file: this.file,
                    resolution: `installDev ${dep}@${project.packageJson.dependencies[dep]}`,
                    position: this.getPositionFromNode(node)
                }], findings);
        });
    }
}
//# sourceMappingURL=FN021005_PKG_types_installed_as_devdep.js.map