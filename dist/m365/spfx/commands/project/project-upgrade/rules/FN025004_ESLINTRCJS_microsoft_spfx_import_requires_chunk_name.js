import ts from 'typescript';
import { TsRule } from './TsRule.js';
export class FN025004_ESLINTRCJS_microsoft_spfx_import_requires_chunk_name extends TsRule {
    constructor() {
        super();
    }
    get id() {
        return 'FN025004';
    }
    get title() {
        return '.eslintrc.js override rule @microsoft/spfx/import-requires-chunk-name';
    }
    get description() {
        return `Remove override rule @microsoft/spfx/import-requires-chunk-name in .eslintrc.js`;
    }
    get resolution() {
        return '';
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
        const node = nodes
            .filter(node => ts.isPropertyName(node))
            .map(node => node)
            .find(i => i.getText() === `'@microsoft/spfx/import-requires-chunk-name'`);
        if (!node) {
            return;
        }
        this.addOccurrence(this.resolution, project.esLintRcJs.path, project.path, node, occurrences);
        if (occurrences.length > 0) {
            this.addFindingWithOccurrences(occurrences, findings);
        }
    }
}
//# sourceMappingURL=FN025004_ESLINTRCJS_microsoft_spfx_import_requires_chunk_name.js.map