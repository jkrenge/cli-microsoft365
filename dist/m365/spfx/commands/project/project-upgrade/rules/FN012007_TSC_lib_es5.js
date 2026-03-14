import { JsonRule } from '../../JsonRule.js';
export class FN012007_TSC_lib_es5 extends JsonRule {
    constructor() {
        super();
    }
    get id() {
        return 'FN012007';
    }
    get title() {
        return 'tsconfig.json es5 lib';
    }
    get description() {
        return `Add es5 lib in tsconfig.json`;
    }
    get resolution() {
        return `{
  "compilerOptions": {
    "lib": [
      "es5"
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
        if (!project.tsConfigJson.compilerOptions.lib ||
            project.tsConfigJson.compilerOptions.lib.indexOf('es5') < 0) {
            const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.lib');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN012007_TSC_lib_es5.js.map