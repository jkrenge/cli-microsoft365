import { JsonRule } from '../../JsonRule.js';
export class FN021007_PKG_only_one_rush_stack_compiler_installed extends JsonRule {
    get id() {
        return 'FN021007';
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
        if (!project.packageJson ||
            !project.packageJson.dependencies) {
            return;
        }
        const projectDeps = Object.keys(project.packageJson.dependencies);
        const projectDevDeps = [];
        if (project.packageJson.devDependencies) {
            // we need to keep them separate so that in case of an error we can
            // determine if the particular dep is a dev dep or not to return correct
            // node
            projectDevDeps.push(...Object.keys(project.packageJson.devDependencies));
            projectDeps.push(...projectDevDeps);
        }
        const rushStacks = projectDeps.filter(dep => dep.indexOf('@microsoft/rush-stack-compiler') === 0);
        if (rushStacks.length <= 1) {
            return;
        }
        const rushStackInTsConfig = this.getRushStackVersionFromTsConfig(project.tsConfigJson?.extends) ?? rushStacks[0];
        rushStacks.forEach(rushStack => {
            if (rushStack === rushStackInTsConfig) {
                return;
            }
            const isDevDep = projectDevDeps.includes(rushStack);
            const node = this.getAstNodeFromFile(project.packageJson, `${isDevDep ? 'devDependencies' : 'dependencies'}.${rushStack}`);
            this.addFindingWithCustomInfo(`Obsolete ${rushStack} found in ${isDevDep ? 'devDependencies' : 'dependencies'}`, `Multiple rush-stack-compiler packages found. Uninstall obsolete ${rushStack}`, [{
                    file: this.file,
                    resolution: `${isDevDep ? 'uninstallDev' : 'uninstall'} ${rushStack}`,
                    position: this.getPositionFromNode(node)
                }], findings);
        });
    }
    getRushStackVersionFromTsConfig(tsConfigExtends) {
        if (!tsConfigExtends) {
            return;
        }
        const match = /@microsoft\/rush-stack-compiler[^/]+/.exec(tsConfigExtends);
        if (!match) {
            return;
        }
        return match[0];
    }
}
//# sourceMappingURL=FN021007_PKG_only_one_rush_stack_compiler_installed.js.map