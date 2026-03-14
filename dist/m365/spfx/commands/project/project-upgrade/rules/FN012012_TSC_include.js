import { JsonRule } from '../../JsonRule.js';
export class FN012012_TSC_include extends JsonRule {
    constructor(include) {
        super();
        this.include = include;
    }
    get id() {
        return 'FN012012';
    }
    get title() {
        return 'tsconfig.json include property';
    }
    get description() {
        return `Add to the tsconfig.json include property`;
    }
    get resolution() {
        return JSON.stringify({
            include: this.include
        }, null, 2);
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
        if (!project.tsConfigJson) {
            return;
        }
        if (!project.tsConfigJson.include ||
            this.include.filter(i => project.tsConfigJson.include.indexOf(i) < 0).length > 0) {
            const node = this.getAstNodeFromFile(project.tsConfigJson, 'include');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN012012_TSC_include.js.map