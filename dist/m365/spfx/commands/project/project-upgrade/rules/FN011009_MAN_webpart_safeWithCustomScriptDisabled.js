import { ManifestRule } from "./ManifestRule.js";
export class FN011009_MAN_webpart_safeWithCustomScriptDisabled extends ManifestRule {
    constructor(add) {
        super();
        this.add = add;
    }
    get id() {
        return 'FN011009';
    }
    get title() {
        return 'Web part manifest safeWithCustomScriptDisabled';
    }
    get description() {
        return `${this.add ? 'Update' : 'Remove'} the safeWithCustomScriptDisabled property in the manifest`;
    }
    get resolution() {
        return `{
  "safeWithCustomScriptDisabled": false
}`;
    }
    get severity() {
        return 'Required';
    }
    visit(project, findings) {
        if (!project.manifests ||
            project.manifests.length === 0) {
            return;
        }
        const occurrences = [];
        project.manifests.forEach(manifest => {
            if (manifest.componentType !== 'WebPart') {
                return;
            }
            if ((this.add && manifest.safeWithCustomScriptDisabled === undefined) ||
                (!this.add && manifest.safeWithCustomScriptDisabled !== undefined)) {
                const node = this.getAstNodeFromFile(manifest, 'safeWithCustomScriptDisabled');
                this.addOccurrence(this.resolution, manifest.path, project.path, node, occurrences);
            }
        });
        if (occurrences.length > 0) {
            this.addFindingWithOccurrences(occurrences, findings);
        }
    }
}
//# sourceMappingURL=FN011009_MAN_webpart_safeWithCustomScriptDisabled.js.map