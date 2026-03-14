import { JsonRule } from '../../JsonRule.js';
export class FN012018_TSC_lib_es2015_promise extends JsonRule {
    constructor() {
        super();
    }
    get id() {
        return 'FN012018';
    }
    get title() {
        return 'tsconfig.json es2015.promise lib';
    }
    get description() {
        return `Add es2015.promise lib in tsconfig.json`;
    }
    get resolution() {
        return `{
  "compilerOptions": {
    "lib": [
      "es2015.promise"
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
            project.tsConfigJson.compilerOptions.lib.indexOf('es2015.promise') < 0) {
            const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.lib');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN012018_TSC_lib_es2015_promise.js.map