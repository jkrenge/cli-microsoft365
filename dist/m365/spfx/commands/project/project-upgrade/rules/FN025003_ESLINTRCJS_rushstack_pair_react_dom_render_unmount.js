import os from 'os';
import ts from 'typescript';
import { TsRule } from './TsRule.js';
export class FN025003_ESLINTRCJS_rushstack_pair_react_dom_render_unmount extends TsRule {
    constructor() {
        super();
    }
    get id() {
        return 'FN025003';
    }
    get title() {
        return '.eslintrc.js override rule @rushstack/pair-react-dom-render-unmount';
    }
    get description() {
        return `Add override rule @rushstack/pair-react-dom-render-unmount in .eslintrc.js`;
    }
    get resolution() {
        return `// Ensure that React components rendered with ReactDOM.render() are unmounted with ReactDOM.unmountComponentAtNode(). https://www.npmjs.com/package/@rushstack/eslint-plugin
        '@rushstack/pair-react-dom-render-unmount': 1,${os.EOL}`;
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
            .filter(i => i.getText() === `'@rushstack/pair-react-dom-render-unmount'`).length !== 0) {
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
//# sourceMappingURL=FN025003_ESLINTRCJS_rushstack_pair_react_dom_render_unmount.js.map