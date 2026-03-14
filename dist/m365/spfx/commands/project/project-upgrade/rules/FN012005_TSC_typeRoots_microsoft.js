import { JsonRule } from '../../JsonRule.js';
export class FN012005_TSC_typeRoots_microsoft extends JsonRule {
    constructor() {
        super();
    }
    get id() {
        return 'FN012005';
    }
    get title() {
        return 'tsconfig.json typeRoots ./node_modules/@microsoft';
    }
    get description() {
        return `Add ./node_modules/@microsoft to typeRoots in tsconfig.json`;
    }
    get resolution() {
        return `{
  "compilerOptions": {
    "typeRoots": [
      "./node_modules/@microsoft"
    ]
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
        if (!project.tsConfigJson.compilerOptions.typeRoots ||
            project.tsConfigJson.compilerOptions.typeRoots.indexOf('./node_modules/@microsoft') < 0) {
            const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.typeRoots');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN012005_TSC_typeRoots_microsoft.js.map