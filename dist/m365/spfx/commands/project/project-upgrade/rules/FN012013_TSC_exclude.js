import { JsonRule } from '../../JsonRule.js';
export class FN012013_TSC_exclude extends JsonRule {
    constructor(exclude, add = true) {
        super();
        this.exclude = exclude;
        this.add = add;
    }
    get id() {
        return 'FN012013';
    }
    get title() {
        return 'tsconfig.json exclude property';
    }
    get description() {
        if (this.add) {
            return `Update tsconfig.json exclude property`;
        }
        else {
            return `Remove tsconfig.json exclude property`;
        }
    }
    get resolution() {
        return JSON.stringify({
            exclude: this.exclude
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
        if (this.add) {
            if (!project.tsConfigJson.exclude ||
                this.exclude.filter(e => project.tsConfigJson.exclude.indexOf(e) < 0).length > 0) {
                const node = this.getAstNodeFromFile(project.tsConfigJson, 'exclude');
                this.addFindingWithPosition(findings, node);
            }
        }
        else {
            if (project.tsConfigJson.exclude) {
                const node = this.getAstNodeFromFile(project.tsConfigJson, 'exclude');
                this.addFindingWithPosition(findings, node);
            }
        }
    }
}
//# sourceMappingURL=FN012013_TSC_exclude.js.map