import { JsonRule } from '../../JsonRule.js';
export class FN014004_CODE_settings_jsonSchemas_configJson_url extends JsonRule {
    constructor(url) {
        super();
        this.url = url;
    }
    get id() {
        return 'FN014004';
    }
    get title() {
        return 'URL of the config.json JSON schema in .vscode/settings.json';
    }
    get description() {
        return `Update the URL of the config.json JSON schema in .vscode/settings.json`;
    }
    get resolution() {
        return `{
  "json.schemas": [
    {
      "fileMatch": [
        "/config/config.json"
      ],
      "url": "./node_modules/@microsoft/sp-build-core-tasks/lib/configJson/schemas/config-v1.schema.json"
    }
  ]
}`;
    }
    get resolutionType() {
        return 'json';
    }
    get severity() {
        return 'Required';
    }
    get file() {
        return '.vscode/settings.json';
    }
    visit(project, findings) {
        if (!project.vsCode ||
            !project.vsCode.settingsJson ||
            !project.vsCode.settingsJson["json.schemas"]) {
            return;
        }
        const schemas = project.vsCode.settingsJson["json.schemas"];
        for (let i = 0; i < schemas.length; i++) {
            const schema = schemas[i];
            if (schema.fileMatch.indexOf('/config/config.json') === -1) {
                continue;
            }
            if (schema.url !== this.url) {
                const node = this.getAstNodeFromFile(project.vsCode.settingsJson, `json;#schemas[${i}]`);
                this.addFindingWithPosition(findings, node);
            }
            return;
        }
    }
}
//# sourceMappingURL=FN014004_CODE_settings_jsonSchemas_configJson_url.js.map