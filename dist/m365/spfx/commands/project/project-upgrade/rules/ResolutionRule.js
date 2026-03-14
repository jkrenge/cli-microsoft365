import { JsonRule } from '../../JsonRule.js';
export class ResolutionRule extends JsonRule {
    constructor(packageName, packageVersion) {
        super();
        this.packageName = packageName;
        this.packageVersion = packageVersion;
    }
    get title() {
        return this.packageName;
    }
    get description() {
        return `Add resolution for package ${this.packageName}`;
    }
    get resolution() {
        return `{
  "resolutions": {
    "${this.packageName}": "${this.packageVersion}"
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
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    customCondition(project) {
        return false;
    }
    visit(project, findings) {
        if (!project.packageJson) {
            return;
        }
        if (this.customCondition(project) &&
            (!project.packageJson.resolutions ||
                project.packageJson.resolutions[this.packageName] !== this.packageVersion)) {
            const node = this.getAstNodeFromFile(project.packageJson, `resolutions.${this.packageName}`);
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=ResolutionRule.js.map