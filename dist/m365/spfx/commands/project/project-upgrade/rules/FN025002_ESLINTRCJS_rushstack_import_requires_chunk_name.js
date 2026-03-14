import os from 'os';
import ts from 'typescript';
import { TsRule } from './TsRule.js';
export class FN025002_ESLINTRCJS_rushstack_import_requires_chunk_name extends TsRule {
    constructor() {
        super();
    }
    get id() {
        return 'FN025002';
    }
    get title() {
        return '.eslintrc.js override rule @rushstack/import-requires-chunk-name';
    }
    get description() {
        return `Add override rule @rushstack/import-requires-chunk-name in .eslintrc.js`;
    }
    get resolution() {
        return `// Require chunk names for dynamic imports in SPFx projects. https://www.npmjs.com/package/@rushstack/eslint-plugin
        '@rushstack/import-requires-chunk-name': 1,${os.EOL}`;
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
            .filter(node => ts.isPropertyName(node))
            .map(node => node)
            .filter(i => i.getText() === `'@rushstack/import-requires-chunk-name'`).length !== 0) {
            return;
        }
        const node = nodes
            .map(node => node)
            .find(i => i.text === 'rules');
        if (!node) {
            return;
        }
        this.addOccurrence(this.resolution, project.esLintRcJs.path, project.path, node, occurrences);
        if (occurrences.length > 0) {
            this.addFindingWithOccurrences(occurrences, findings);
        }
    }
}
//# sourceMappingURL=FN025002_ESLINTRCJS_rushstack_import_requires_chunk_name.js.map