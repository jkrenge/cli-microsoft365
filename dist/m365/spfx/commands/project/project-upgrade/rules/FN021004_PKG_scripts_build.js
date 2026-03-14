import { JsonRule } from "../../JsonRule.js";
export class FN021004_PKG_scripts_build extends JsonRule {
    constructor(script) {
        super();
        this.script = script;
    }
    get id() {
        return 'FN021004';
    }
    get title() {
        return 'package.json scripts.build';
    }
    get description() {
        return 'Update package.json scripts.build property';
    }
    get resolution() {
        return `{
  "scripts": {
    "build": "${this.script}"
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
            !project.packageJson.scripts.build ||
            project.packageJson.scripts.build !== this.script) {
            const node = this.getAstNodeFromFile(project.packageJson, 'scripts.build');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN021004_PKG_scripts_build.js.map