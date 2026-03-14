import { JsonRule } from '../../JsonRule.js';
export class FN012003_TSC_skipLibCheck extends JsonRule {
    constructor(skipLibCheck) {
        super();
        this.skipLibCheck = skipLibCheck;
    }
    get id() {
        return 'FN012003';
    }
    get title() {
        return 'tsconfig.json skipLibCheck';
    }
    get description() {
        return `Update skipLibCheck in tsconfig.json`;
    }
    get resolution() {
        return `{
  "compilerOptions": {
    "skipLibCheck": true
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
        if (project.tsConfigJson.compilerOptions.skipLibCheck !== this.skipLibCheck) {
            const node = this.getAstNodeFromFile(project.tsConfigJson, 'compilerOptions.skipLibCheck');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN012003_TSC_skipLibCheck.js.map