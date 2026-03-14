import { JsonRule } from "../../JsonRule.js";
export class FN021005_PKG_scripts_test extends JsonRule {
    constructor(script, add = true) {
        super();
        this.script = script;
        this.add = add;
    }
    get id() {
        return 'FN021005';
    }
    get title() {
        return 'package.json scripts.test';
    }
    get description() {
        return `${this.add ? 'Update' : 'Remove'} package.json scripts.test property`;
    }
    get resolution() {
        return `{
  "scripts": {
    "test": ${this.add ? `"${this.script}"` : '""'}
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
        return './package.json';
    }
    visit(project, findings) {
        if (!project.packageJson) {
            return;
        }
        if (this.add) {
            if (!project.packageJson.scripts ||
                typeof project.packageJson.scripts !== 'object' ||
                !project.packageJson.scripts.test ||
                project.packageJson.scripts.test !== this.script) {
                const node = this.getAstNodeFromFile(project.packageJson, 'scripts.test');
                this.addFindingWithPosition(findings, node);
            }
        }
        else {
            if (project.packageJson.scripts?.test === this.script) {
                const node = this.getAstNodeFromFile(project.packageJson, 'scripts.test');
                this.addFindingWithPosition(findings, node);
            }
        }
    }
}
//# sourceMappingURL=FN021005_PKG_scripts_test.js.map