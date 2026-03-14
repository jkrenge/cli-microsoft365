import { JsonRule } from '../../JsonRule.js';
export class FN012019_TSC_types_es6_promise extends JsonRule {
    constructor(add) {
        super();
        this.add = add;
    }
    get id() {
        return 'FN012019';
    }
    get title() {
        return 'tsconfig.json es6-promise types';
    }
    get description() {
        return `${(this.add ? 'Add' : 'Remove')} es6-promise type in tsconfig.json`;
    }
    get resolution() {
        return `{
  "compilerOptions": {
    "types": [
      "es6-promise"
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
                project.tsConfigJson.compilerOptions.types.indexOf('es6-promise') < 0) {
                const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.types');
                this.addFindingWithPosition(findings, node);
            }
        }
        else {
            if (project.tsConfigJson.compilerOptions.types &&
                project.tsConfigJson.compilerOptions.types.indexOf('es6-promise') > -1) {
                const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.types[es6-promise]');
                this.addFindingWithPosition(findings, node);
            }
        }
    }
}
//# sourceMappingURL=FN012019_TSC_types_es6_promise.js.map