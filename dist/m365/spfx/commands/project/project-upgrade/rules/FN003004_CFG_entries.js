import { JsonRule } from '../../JsonRule.js';
export class FN003004_CFG_entries extends JsonRule {
    get id() {
        return 'FN003004';
    }
    get title() {
        return `config.json entries`;
    }
    get description() {
        return `Remove the "entries" property in ${this.file}`;
    }
    get resolution() {
        return '';
    }
    get resolutionType() {
        return 'json';
    }
    get severity() {
        return 'Required';
    }
    get file() {
        return './config/config.json';
    }
    visit(project, findings) {
        if (!project.configJson) {
            return;
        }
        const entries = project.configJson.entries;
        if (entries !== undefined) {
            const node = this.getAstNodeFromFile(project.configJson, 'entries');
            this.addFindingWithOccurrences([{
                    file: this.file,
                    resolution: JSON.stringify({ entries: entries }, null, 2),
                    position: this.getPositionFromNode(node)
                }], findings);
        }
    }
}
//# sourceMappingURL=FN003004_CFG_entries.js.map