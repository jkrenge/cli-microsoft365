import { JsonRule } from '../../JsonRule.js';
export class FN012004_TSC_typeRoots_types extends JsonRule {
    constructor() {
        super();
    }
    get id() {
        return 'FN012004';
    }
    get title() {
        return 'tsconfig.json typeRoots ./node_modules/@types';
    }
    get description() {
        return `Add ./node_modules/@types to typeRoots in tsconfig.json`;
    }
    get resolution() {
        return `{
  "compilerOptions": {
    "typeRoots": [
      "./node_modules/@types"
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
            project.tsConfigJson.compilerOptions.typeRoots.indexOf('./node_modules/@types') < 0) {
            const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.typeRoots');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN012004_TSC_typeRoots_types.js.map