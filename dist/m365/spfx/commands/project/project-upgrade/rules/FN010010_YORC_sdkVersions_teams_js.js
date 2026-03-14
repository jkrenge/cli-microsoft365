import { JsonRule } from '../../JsonRule.js';
export class FN010010_YORC_sdkVersions_teams_js extends JsonRule {
    constructor(version) {
        super();
        this.version = version;
    }
    get id() {
        return 'FN010010';
    }
    get title() {
        return '.yo-rc.json @microsoft/teams-js SDK version';
    }
    get description() {
        return `Update @microsoft/teams-js SDK version in .yo-rc.json`;
    }
    get resolution() {
        return `{
  "@microsoft/generator-sharepoint": {
    "sdkVersions": {
      "@microsoft/teams-js": "${this.version}"
    }
  }
}`;
    }
    get resolutionType() {
        return 'json';
    }
    get severity() {
        return 'Recommended';
    }
    get file() {
        return './.yo-rc.json';
    }
    visit(project, findings) {
        if (!project.yoRcJson) {
            return;
        }
        if (project.yoRcJson["@microsoft/generator-sharepoint"]?.sdkVersions?.['@microsoft/teams-js'] !== this.version) {
            let nodePath = '@microsoft/generator-sharepoint';
            if (project.yoRcJson["@microsoft/generator-sharepoint"]?.sdkVersions) {
                nodePath += '.sdkVersions';
                if (project.yoRcJson["@microsoft/generator-sharepoint"].sdkVersions['@microsoft/teams-js']) {
                    nodePath += '.@microsoft/teams-js';
                }
            }
            const node = this.getAstNodeFromFile(project.yoRcJson, nodePath);
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN010010_YORC_sdkVersions_teams_js.js.map