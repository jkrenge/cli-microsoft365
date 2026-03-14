import { JsonRule } from '../../JsonRule.js';
export class FN008002_CFG_TSL_removeRule extends JsonRule {
    constructor(rule) {
        super();
        this.rule = rule;
    }
    get id() {
        return 'FN008002';
    }
    get title() {
        return `tslint.json ${this.rule} rule`;
    }
    get description() {
        return `In tslint.json remove the ${this.rule} rule`;
    }
    get resolution() {
        return `{
  "lintConfig": {
    "rules": {
      "${this.rule}": false
    }
  }
}`;
    }
    get resolutionType() {
        return 'json';
    }
    get severity() {
        return 'Recommended';
    }
    get file() {
        return './config/tslint.json';
    }
    visit(project, findings) {
        if (!project.tsLintJson ||
            !project.tsLintJson ||
            !project.tsLintJson.lintConfig ||
            !project.tsLintJson.lintConfig.rules) {
            return;
        }
        if (typeof project.tsLintJson.lintConfig.rules[this.rule] !== 'undefined') {
            const node = this.getAstNodeFromFile(project.tsLintJson, `lintConfig.rules.${this.rule}`);
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN008002_CFG_TSL_removeRule.js.map