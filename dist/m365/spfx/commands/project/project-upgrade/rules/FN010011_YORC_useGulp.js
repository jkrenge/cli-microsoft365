import { JsonRule } from '../../JsonRule.js';
export class FN010011_YORC_useGulp extends JsonRule {
    constructor(useGulp) {
        super();
        this.useGulp = useGulp;
    }
    get id() {
        return 'FN010011';
    }
    get title() {
        return '.yo-rc.json useGulp';
    }
    get description() {
        return `Update useGulp property in .yo-rc.json`;
    }
    get resolution() {
        return `{
    "useGulp": ${this.useGulp.toString()}
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
        if (project.yoRcJson["@microsoft/generator-sharepoint"]?.useGulp !== this.useGulp) {
            const node = this.getAstNodeFromFile(project.yoRcJson, '@microsoft/generator-sharepoint.useGulp');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN010011_YORC_useGulp.js.map