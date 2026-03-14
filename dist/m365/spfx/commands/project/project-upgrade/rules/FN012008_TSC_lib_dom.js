import { JsonRule } from '../../JsonRule.js';
export class FN012008_TSC_lib_dom extends JsonRule {
    constructor() {
        super();
    }
    get id() {
        return 'FN012008';
    }
    get title() {
        return 'tsconfig.json dom lib';
    }
    get description() {
        return `Add dom lib in tsconfig.json`;
    }
    get resolution() {
        return `{
  "compilerOptions": {
    "lib": [
      "dom"
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
            project.tsConfigJson.compilerOptions.lib.indexOf('dom') < 0) {
            const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.lib');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN012008_TSC_lib_dom.js.map