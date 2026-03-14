import { JsonRule } from '../../JsonRule.js';
export class FN012015_TSC_strictNullChecks extends JsonRule {
    constructor(strictNullChecks) {
        super();
        this.strictNullChecks = strictNullChecks;
    }
    get id() {
        return 'FN012015';
    }
    get title() {
        return 'tsconfig.json compiler options strictNullChecks';
    }
    get description() {
        return `Update tsconfig.json strictNullChecks value`;
    }
    get resolution() {
        return `{
  "compilerOptions": {
    "strictNullChecks": ${this.strictNullChecks}
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
        return './tsconfig.json';
    }
    visit(project, findings) {
        if (!project.tsConfigJson || !project.tsConfigJson.compilerOptions) {
            return;
        }
        if (project.tsConfigJson.compilerOptions.strictNullChecks !== this.strictNullChecks) {
            const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.strictNullChecks');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN012015_TSC_strictNullChecks.js.map