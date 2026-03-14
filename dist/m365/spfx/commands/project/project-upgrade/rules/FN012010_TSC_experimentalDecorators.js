import { JsonRule } from '../../JsonRule.js';
export class FN012010_TSC_experimentalDecorators extends JsonRule {
    constructor() {
        super();
    }
    get id() {
        return 'FN012010';
    }
    get title() {
        return 'tsconfig.json compiler options experimental decorators';
    }
    get description() {
        return `Enable tsconfig.json experimental decorators`;
    }
    get resolution() {
        return `{
  "compilerOptions": {
    "experimentalDecorators": true
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
        if (!project.tsConfigJson.compilerOptions.experimentalDecorators) {
            const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.experimentalDecorators');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN012010_TSC_experimentalDecorators.js.map