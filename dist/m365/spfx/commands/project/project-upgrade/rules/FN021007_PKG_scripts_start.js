import { JsonRule } from "../../JsonRule.js";
export class FN021007_PKG_scripts_start extends JsonRule {
    constructor(script) {
        super();
        this.script = script;
    }
    get id() {
        return 'FN021007';
    }
    get title() {
        return 'package.json scripts.start';
    }
    get description() {
        return 'Update package.json scripts.start property';
    }
    get resolution() {
        return `{
  "scripts": {
    "start": "${this.script}"
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
            !project.packageJson.scripts.start ||
            project.packageJson.scripts.start !== this.script) {
            const node = this.getAstNodeFromFile(project.packageJson, 'scripts.start');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN021007_PKG_scripts_start.js.map