import { JsonRule } from '../../JsonRule.js';
export class FN026001_CFG_SASS_schema extends JsonRule {
    constructor(schema) {
        super();
        this.schema = schema;
    }
    get id() {
        return 'FN026001';
    }
    get title() {
        return 'sass.json schema';
    }
    get description() {
        return `Update sass.json schema URL`;
    }
    get resolution() {
        return `{
  "$schema": "${this.schema}"
}`;
    }
    get resolutionType() {
        return 'json';
    }
    get severity() {
        return 'Required';
    }
    get file() {
        return './config/sass.json';
    }
    visit(project, findings) {
        if (!project.sassJson) {
            return;
        }
        if (project.sassJson.$schema !== this.schema) {
            const node = this.getAstNodeFromFile(project.sassJson, '$schema');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN026001_CFG_SASS_schema.js.map