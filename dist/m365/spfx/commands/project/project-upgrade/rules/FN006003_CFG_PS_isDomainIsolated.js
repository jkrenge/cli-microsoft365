import { JsonRule } from '../../JsonRule.js';
export class FN006003_CFG_PS_isDomainIsolated extends JsonRule {
    constructor(isDomainIsolated) {
        super();
        this.isDomainIsolated = isDomainIsolated;
    }
    get id() {
        return 'FN006003';
    }
    get title() {
        return 'package-solution.json isDomainIsolated';
    }
    get description() {
        return `Update package-solution.json isDomainIsolated`;
    }
    get resolution() {
        return `{
  "solution": {
    "isDomainIsolated": ${this.isDomainIsolated}
  }
}`;
    }
    get resolutionType() {
        return 'json';
    }
    get severity() {
        return 'Required';
    }
    get file() {
        return './config/package-solution.json';
    }
    visit(project, findings) {
        if (!project.packageSolutionJson ||
            !project.packageSolutionJson.solution) {
            return;
        }
        if (project.packageSolutionJson.solution.isDomainIsolated !== this.isDomainIsolated) {
            const node = this.getAstNodeFromFile(project.packageSolutionJson, 'solution.isDomainIsolated');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN006003_CFG_PS_isDomainIsolated.js.map