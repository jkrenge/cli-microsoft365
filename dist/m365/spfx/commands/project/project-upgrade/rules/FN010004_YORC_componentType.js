import { JsonRule } from '../../JsonRule.js';
export class FN010004_YORC_componentType extends JsonRule {
    constructor() {
        super();
    }
    get id() {
        return 'FN010004';
    }
    get title() {
        return '.yo-rc.json componentType';
    }
    get description() {
        return `Update componentType in .yo-rc.json`;
    }
    get resolution() {
        return '';
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
        let componentType;
        if (project.manifests) {
            for (let i = 0; i < project.manifests.length; i++) {
                const manifest = project.manifests[i];
                if (manifest.componentType === 'WebPart') {
                    componentType = 'webpart';
                    break;
                }
                if (manifest.componentType === 'Extension') {
                    componentType = 'extension';
                    break;
                }
            }
        }
        if (!componentType) {
            componentType = 'webpart';
        }
        if (project.yoRcJson["@microsoft/generator-sharepoint"]?.componentType !== componentType) {
            const node = this.getAstNodeFromFile(project.yoRcJson, '@microsoft/generator-sharepoint.componentType');
            this.addFindingWithOccurrences([{
                    file: this.file,
                    resolution: JSON.stringify({
                        "@microsoft/generator-sharepoint": {
                            "componentType": componentType
                        }
                    }, null, 2),
                    position: this.getPositionFromNode(node)
                }], findings);
        }
    }
}
//# sourceMappingURL=FN010004_YORC_componentType.js.map