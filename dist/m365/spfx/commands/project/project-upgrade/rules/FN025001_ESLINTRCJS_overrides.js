import ts from 'typescript';
import { TsRule } from './TsRule.js';
export class FN025001_ESLINTRCJS_overrides extends TsRule {
    constructor(contents) {
        super();
        this.contents = contents;
    }
    get id() {
        return 'FN025001';
    }
    get title() {
        return '.eslintrc.js overrides';
    }
    get description() {
        return `Add overrides in .eslintrc.js`;
    }
    get resolution() {
        return `export default {
      overrides: [
        ${this.contents}
      ]
    };`;
    }
    get resolutionType() {
        return 'js';
    }
    get severity() {
        return 'Required';
    }
    get file() {
        return './.eslintrc.js';
    }
    visit(project, findings) {
        if (!project.esLintRcJs) {
            return;
        }
        const nodes = project.esLintRcJs.nodes;
        if (!nodes) {
            return;
        }
        const occurrences = [];
        if (nodes
            .filter(node => ts.isIdentifier(node))
            .map(node => node)
            .filter(i => i.getText() === 'overrides').length !== 0) {
            return;
        }
        const node = nodes
            .map(node => node)
            .find(i => i.text === 'module');
        if (!node) {
            return;
        }
        this.addOccurrence(this.resolution, project.esLintRcJs.path, project.path, node, occurrences);
        if (occurrences.length > 0) {
            this.addFindingWithOccurrences(occurrences, findings);
        }
    }
}
//# sourceMappingURL=FN025001_ESLINTRCJS_overrides.js.map