import { JsonRule } from '../../JsonRule.js';
export class FN007002_CFG_S_initialPage extends JsonRule {
    constructor(initialPage) {
        super();
        this.initialPage = initialPage;
    }
    get id() {
        return 'FN007002';
    }
    get title() {
        return 'serve.json initialPage';
    }
    get description() {
        return `Update serve.json initialPage URL`;
    }
    get resolution() {
        return `{
  "initialPage": "${this.initialPage}"
}`;
    }
    get resolutionType() {
        return 'json';
    }
    get severity() {
        return 'Required';
    }
    get file() {
        return './config/serve.json';
    }
    visit(project, findings) {
        if (!project.serveJson) {
            return;
        }
        if (project.serveJson.initialPage !== this.initialPage) {
            const node = this.getAstNodeFromFile(project.serveJson, 'initialPage');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN007002_CFG_S_initialPage.js.map