import { JsonRule } from '../../JsonRule.js';
export class FN012001_TSC_module extends JsonRule {
    constructor(module) {
        super();
        this.module = module;
    }
    get id() {
        return 'FN012001';
    }
    get title() {
        return 'tsconfig.json module';
    }
    get description() {
        return `Update module type in tsconfig.json`;
    }
    get resolution() {
        return `{
  "compilerOptions": {
    "module": "${this.module}"
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
        if (project.tsConfigJson.compilerOptions.module !== this.module) {
            const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.module');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN012001_TSC_module.js.map