import { JsonRule } from '../../JsonRule.js';
export class PackageRule extends JsonRule {
    constructor(propertyName, add, propertyValue) {
        super();
        this.propertyName = propertyName;
        this.add = add;
        this.propertyValue = propertyValue;
    }
    get title() {
        return this.propertyName;
    }
    get description() {
        return `${this.add ? 'Add' : 'Remove'} package.json property`;
    }
    get resolution() {
        return `{
  "${this.propertyName}": "${this.propertyValue}"
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
        if (!project.packageJson ||
            (this.add && !project.packageJson[this.propertyName]) ||
            (!this.add && project.packageJson[this.propertyName])) {
            const node = this.getAstNodeFromFile(project.packageJson, this.propertyName);
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=PackageRule.js.map