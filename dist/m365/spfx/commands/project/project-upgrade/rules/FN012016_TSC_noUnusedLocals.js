import { JsonRule } from '../../JsonRule.js';
export class FN012016_TSC_noUnusedLocals extends JsonRule {
    constructor(noUnusedLocals) {
        super();
        this.noUnusedLocals = noUnusedLocals;
    }
    get id() {
        return 'FN012016';
    }
    get title() {
        return 'tsconfig.json compiler options noUnusedLocals';
    }
    get description() {
        return `Update tsconfig.json noUnusedLocals value`;
    }
    get resolution() {
        return `{
  "compilerOptions": {
    "noUnusedLocals": ${this.noUnusedLocals}
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
        if (project.tsConfigJson.compilerOptions.noUnusedLocals !== this.noUnusedLocals) {
            const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.noUnusedLocals');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN012016_TSC_noUnusedLocals.js.map