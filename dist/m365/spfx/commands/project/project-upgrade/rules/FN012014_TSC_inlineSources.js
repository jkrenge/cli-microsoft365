import { JsonRule } from '../../JsonRule.js';
export class FN012014_TSC_inlineSources extends JsonRule {
    constructor(inlineSources) {
        super();
        this.inlineSources = inlineSources;
    }
    get id() {
        return 'FN012014';
    }
    get title() {
        return 'tsconfig.json compiler options inlineSources';
    }
    get description() {
        return `Update tsconfig.json inlineSources value`;
    }
    get resolution() {
        return `{
  "compilerOptions": {
    "inlineSources": ${this.inlineSources}
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
        if (project.tsConfigJson.compilerOptions.inlineSources !== this.inlineSources) {
            const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.inlineSources');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN012014_TSC_inlineSources.js.map