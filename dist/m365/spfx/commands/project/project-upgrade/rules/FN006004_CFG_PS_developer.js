import { JsonRule } from '../../JsonRule.js';
export class FN006004_CFG_PS_developer extends JsonRule {
    constructor(version) {
        super();
        this.version = version;
    }
    get id() {
        return 'FN006004';
    }
    get title() {
        return 'package-solution.json developer';
    }
    get description() {
        return `In package-solution.json add developer section`;
    }
    get resolution() {
        return `{
  "solution": {
    "developer": {
      "name": "",
      "privacyUrl": "",
      "termsOfUseUrl": "",
      "websiteUrl": "",
      "mpnId": "${this.version ? `Undefined-${this.version}` : ''}"
    }
  }
}`;
    }
    get resolutionType() {
        return 'json';
    }
    get severity() {
        return 'Optional';
    }
    get file() {
        return './config/package-solution.json';
    }
    visit(project, findings) {
        if (!project.packageSolutionJson ||
            !project.packageSolutionJson.solution) {
            return;
        }
        if (!project.packageSolutionJson.solution.developer) {
            const node = this.getAstNodeFromFile(project.packageSolutionJson, 'solution');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN006004_CFG_PS_developer.js.map