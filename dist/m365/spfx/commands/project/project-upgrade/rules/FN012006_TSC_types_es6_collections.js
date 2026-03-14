import { JsonRule } from '../../JsonRule.js';
export class FN012006_TSC_types_es6_collections extends JsonRule {
    constructor(add) {
        super();
        this.add = add;
    }
    get id() {
        return 'FN012006';
    }
    get title() {
        return 'tsconfig.json es6-collections types';
    }
    get description() {
        return `${(this.add ? 'Add' : 'Remove')} es6-collections type in tsconfig.json`;
    }
    get resolution() {
        return `{
  "compilerOptions": {
    "types": [
      "es6-collections"
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
        if (this.add) {
            if (!project.tsConfigJson.compilerOptions.types ||
                project.tsConfigJson.compilerOptions.types.indexOf('es6-collections') < 0) {
                const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.types');
                this.addFindingWithPosition(findings, node);
            }
        }
        else {
            if (project.tsConfigJson.compilerOptions.types &&
                project.tsConfigJson.compilerOptions.types.indexOf('es6-collections') > -1) {
                const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.types[es6-collections]');
                this.addFindingWithPosition(findings, node);
            }
        }
    }
}
//# sourceMappingURL=FN012006_TSC_types_es6_collections.js.map