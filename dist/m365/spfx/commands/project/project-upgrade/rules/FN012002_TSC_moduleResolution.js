import { JsonRule } from '../../JsonRule.js';
export class FN012002_TSC_moduleResolution extends JsonRule {
    constructor(moduleResolution) {
        super();
        this.moduleResolution = moduleResolution;
    }
    get id() {
        return 'FN012002';
    }
    get title() {
        return 'tsconfig.json moduleResolution';
    }
    get description() {
        return `Update moduleResolution in tsconfig.json`;
    }
    get resolution() {
        return `{
  "compilerOptions": {
    "moduleResolution": "${this.moduleResolution}"
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
        if (project.tsConfigJson.compilerOptions.moduleResolution !== this.moduleResolution) {
            const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.moduleResolution');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN012002_TSC_moduleResolution.js.map