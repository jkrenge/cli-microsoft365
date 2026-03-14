import { JsonRule } from '../../JsonRule.js';
export class FN012020_TSC_noImplicitAny extends JsonRule {
    constructor(noImplicitAny) {
        super();
        this.noImplicitAny = noImplicitAny;
    }
    get id() {
        return 'FN012020';
    }
    get title() {
        return 'tsconfig.json noImplicitAny';
    }
    get description() {
        return `Add noImplicitAny in tsconfig.json`;
    }
    get resolution() {
        return `{
  "compilerOptions": {
    "noImplicitAny": ${this.noImplicitAny}
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
        if (project.tsConfigJson.compilerOptions.noImplicitAny !== this.noImplicitAny) {
            const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.noImplicitAny');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN012020_TSC_noImplicitAny.js.map