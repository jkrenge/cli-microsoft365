import { JsonRule } from '../../JsonRule.js';
export class FN010009_YORC_sdkVersions_microsoft_graph_client extends JsonRule {
    constructor(version) {
        super();
        this.version = version;
    }
    get id() {
        return 'FN010009';
    }
    get title() {
        return '.yo-rc.json @microsoft/microsoft-graph-client SDK version';
    }
    get description() {
        return `Update @microsoft/microsoft-graph-client SDK version in .yo-rc.json`;
    }
    get resolution() {
        return `{
  "@microsoft/generator-sharepoint": {
    "sdkVersions": {
      "@microsoft/microsoft-graph-client": "${this.version}"
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
        if (project.yoRcJson["@microsoft/generator-sharepoint"]?.sdkVersions?.['@microsoft/microsoft-graph-client'] !== this.version) {
            let nodePath = '@microsoft/generator-sharepoint';
            if (project.yoRcJson["@microsoft/generator-sharepoint"]?.sdkVersions) {
                nodePath += '.sdkVersions';
                if (project.yoRcJson["@microsoft/generator-sharepoint"].sdkVersions['@microsoft/microsoft-graph-client']) {
                    nodePath += '.@microsoft/microsoft-graph-client';
                }
            }
            const node = this.getAstNodeFromFile(project.yoRcJson, nodePath);
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN010009_YORC_sdkVersions_microsoft_graph_client.js.map