import { JsonRule } from '../../JsonRule.js';
export class FN006002_CFG_PS_includeClientSideAssets extends JsonRule {
    constructor(includeClientSideAssets) {
        super();
        this.includeClientSideAssets = includeClientSideAssets;
    }
    get id() {
        return 'FN006002';
    }
    get title() {
        return 'package-solution.json includeClientSideAssets';
    }
    get description() {
        return `Update package-solution.json includeClientSideAssets`;
    }
    get resolution() {
        return `{
  "solution": {
    "includeClientSideAssets": ${this.includeClientSideAssets}
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
        if (project.packageSolutionJson.solution.includeClientSideAssets !== this.includeClientSideAssets) {
            const node = this.getAstNodeFromFile(project.packageSolutionJson, 'solution.includeClientSideAssets');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN006002_CFG_PS_includeClientSideAssets.js.map