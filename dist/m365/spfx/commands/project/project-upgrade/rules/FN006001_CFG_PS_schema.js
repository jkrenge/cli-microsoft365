import { JsonRule } from '../../JsonRule.js';
export class FN006001_CFG_PS_schema extends JsonRule {
    constructor(schema) {
        super();
        this.schema = schema;
    }
    get id() {
        return 'FN006001';
    }
    get title() {
        return 'package-solution.json schema';
    }
    get description() {
        return `Update package-solution.json schema URL`;
    }
    get resolution() {
        return `{
  "$schema": "${this.schema}"
}`;
    }
    get resolutionType() {
        return 'json';
    }
    get severity() {
        return 'Required';
    }
    get file() {
        return './config/package-solution.json';
    }
    visit(project, findings) {
        if (!project.packageSolutionJson) {
            return;
        }
        if (project.packageSolutionJson.$schema !== this.schema) {
            const node = this.getAstNodeFromFile(project.packageSolutionJson, '$schema');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN006001_CFG_PS_schema.js.map