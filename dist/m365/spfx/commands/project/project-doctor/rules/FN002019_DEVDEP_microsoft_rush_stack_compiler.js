import { JsonRule } from '../../JsonRule.js';
export class FN002019_DEVDEP_microsoft_rush_stack_compiler extends JsonRule {
    constructor(supportedVersions) {
        super();
        this.supportedVersions = supportedVersions;
    }
    get id() {
        return 'FN002019';
    }
    get title() {
        return `@microsoft/rush-stack-compiler-${this.supportedVersions[this.supportedVersions.length - 1]}`;
    }
    get description() {
        return '';
    }
    get resolution() {
        return '';
    }
    get resolutionType() {
        return 'cmd';
    }
    get severity() {
        return 'Required';
    }
    get file() {
        return './package.json';
    }
    visit(project, findings) {
        if (!project.packageJson) {
            return;
        }
        const projectDeps = [];
        if (project.packageJson.dependencies) {
            projectDeps.push(...Object.keys(project.packageJson.dependencies));
        }
        const projectDevDeps = [];
        if (project.packageJson.devDependencies) {
            // we need to keep them separate so that in case of an error we can
            // determine if the particular dep is a dev dep or not to return correct
            // node
            projectDevDeps.push(...Object.keys(project.packageJson.devDependencies));
            projectDeps.push(...projectDevDeps);
        }
        const rushStacks = projectDeps.filter(dep => dep.indexOf('@microsoft/rush-stack-compiler') === 0);
        const occurrences = [];
        rushStacks.forEach(rushStack => {
            const version = rushStack.replace('@microsoft/rush-stack-compiler-', '');
            if (this.supportedVersions.includes(version)) {
                return;
            }
            const node = this.getAstNodeFromFile(project.packageJson, `${projectDevDeps.includes(rushStack) ? 'devDependencies' : 'dependencies'}.${rushStack}`);
            occurrences.push({
                file: this.file,
                resolution: `uninstallDev ${rushStack}`,
                position: this.getPositionFromNode(node)
            });
        });
        if (occurrences.length === 0) {
            return;
        }
        this.addFindingWithCustomInfo(this.title, `Uninstall unsupported version${occurrences.length === 1 ? '' : 's'} of @microsoft/rush-stack-compiler`, occurrences, findings);
    }
}
//# sourceMappingURL=FN002019_DEVDEP_microsoft_rush_stack_compiler.js.map