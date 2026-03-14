import { JsonRule } from "../../JsonRule.js";
export class FN021008_PKG_scripts_eject_webpack extends JsonRule {
    constructor(script) {
        super();
        this.script = script;
    }
    get id() {
        return 'FN021008';
    }
    get title() {
        return 'package.json scripts.eject-webpack';
    }
    get description() {
        return 'Update package.json scripts.eject-webpack property';
    }
    get resolution() {
        return `{
  "scripts": {
    "eject-webpack": "${this.script}"
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
            !project.packageJson.scripts['eject-webpack'] ||
            project.packageJson.scripts['eject-webpack'] !== this.script) {
            const node = this.getAstNodeFromFile(project.packageJson, 'scripts.eject-webpack');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN021008_PKG_scripts_eject_webpack.js.map