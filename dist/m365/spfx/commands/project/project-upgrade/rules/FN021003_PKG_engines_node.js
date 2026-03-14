import { JsonRule } from "../../JsonRule.js";
export class FN021003_PKG_engines_node extends JsonRule {
    constructor(version) {
        super();
        this.version = version;
    }
    get id() {
        return 'FN021003';
    }
    get title() {
        return 'package.json engines.node';
    }
    get description() {
        return 'Update package.json engines.node property';
    }
    get resolution() {
        return `{
  "engines": {
    "node": "${this.version}"
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
        if (!project.packageJson.engines ||
            typeof project.packageJson.engines !== 'object' ||
            !project.packageJson.engines.node ||
            project.packageJson.engines.node !== this.version) {
            const node = this.getAstNodeFromFile(project.packageJson, 'engines.node');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN021003_PKG_engines_node.js.map