import { JsonRule } from "../../JsonRule.js";
export class FN021006_PKG_scripts_clean extends JsonRule {
    constructor(script) {
        super();
        this.script = script;
    }
    get id() {
        return 'FN021006';
    }
    get title() {
        return 'package.json scripts.clean';
    }
    get description() {
        return 'Update package.json scripts.clean property';
    }
    get resolution() {
        return `{
  "scripts": {
    "clean": "${this.script}"
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
        if (!project.packageJson.scripts ||
            typeof project.packageJson.scripts !== 'object' ||
            !project.packageJson.scripts.clean ||
            project.packageJson.scripts.clean !== this.script) {
            const node = this.getAstNodeFromFile(project.packageJson, 'scripts.clean');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN021006_PKG_scripts_clean.js.map