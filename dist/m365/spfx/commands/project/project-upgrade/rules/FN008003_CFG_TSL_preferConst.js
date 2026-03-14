import { JsonRule } from '../../JsonRule.js';
export class FN008003_CFG_TSL_preferConst extends JsonRule {
    constructor() {
        super();
    }
    get id() {
        return 'FN008003';
    }
    get title() {
        return 'tslint.json prefer-const';
    }
    get description() {
        return `Remove prefer-const from tslint.json`;
    }
    get resolution() {
        return `{
  "lintConfig": {
    "rules": {
      "prefer-const": true
    }
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
        return './config/tslint.json';
    }
    visit(project, findings) {
        if (!project.tsLintJson) {
            return;
        }
        if (project.tsLintJson.lintConfig && project.tsLintJson.lintConfig.rules &&
            project.tsLintJson.lintConfig.rules["prefer-const"] === true) {
            const node = this.getAstNodeFromFile(project.tsLintJson, 'lintConfig.rules.prefer-const');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN008003_CFG_TSL_preferConst.js.map