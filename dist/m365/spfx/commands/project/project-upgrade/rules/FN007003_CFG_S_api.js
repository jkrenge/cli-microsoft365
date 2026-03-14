import { JsonRule } from '../../JsonRule.js';
export class FN007003_CFG_S_api extends JsonRule {
    constructor() {
        super();
    }
    get id() {
        return 'FN007003';
    }
    get title() {
        return 'serve.json api';
    }
    get description() {
        return `From serve.json remove the api property`;
    }
    get resolution() {
        return ``;
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
        if (project.serveJson.api) {
            const node = this.getAstNodeFromFile(project.serveJson, 'api');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN007003_CFG_S_api.js.map