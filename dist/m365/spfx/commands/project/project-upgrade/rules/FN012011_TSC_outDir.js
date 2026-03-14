import { JsonRule } from '../../JsonRule.js';
export class FN012011_TSC_outDir extends JsonRule {
    constructor(outDir) {
        super();
        this.outDir = outDir;
    }
    get id() {
        return 'FN012011';
    }
    get title() {
        return 'tsconfig.json compiler options outDir';
    }
    get description() {
        return `Update tsconfig.json outDir value`;
    }
    get resolution() {
        return `{
  "compilerOptions": {
    "outDir": "${this.outDir}"
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
        if (project.tsConfigJson.compilerOptions.outDir !== this.outDir) {
            const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.outDir');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN012011_TSC_outDir.js.map