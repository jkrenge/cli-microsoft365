import { JsonRule } from "../../JsonRule.js";
export class FN021009_PKG_overrides_rushstack_heft extends JsonRule {
    constructor(version) {
        super();
        this.version = version;
    }
    get id() {
        return 'FN021009';
    }
    get title() {
        return 'package.json overrides.@rushstack/heft';
    }
    get description() {
        return 'Update package.json overrides.@rushstack/heft property';
    }
    get resolution() {
        return `{
  "overrides": {
    "@rushstack/heft": "${this.version}"
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
        if (!project.packageJson.overrides ||
            typeof project.packageJson.overrides !== 'object' ||
            !project.packageJson.overrides['@rushstack/heft'] ||
            project.packageJson.overrides['@rushstack/heft'] !== this.version) {
            const node = this.getAstNodeFromFile(project.packageJson, 'overrides.@rushstack/heft');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN021009_PKG_overrides_rushstack_heft.js.map