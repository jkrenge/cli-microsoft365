import { ManifestRule } from "./ManifestRule.js";
export class FN011011_MAN_webpart_supportedHosts extends ManifestRule {
    constructor(add) {
        super();
        this.add = add;
    }
    get id() {
        return 'FN011011';
    }
    get title() {
        return 'Web part manifest supportedHosts';
    }
    get description() {
        return `${this.add ? 'Update' : 'Remove'} the supportedHosts property in the manifest`;
    }
    get resolution() {
        return `{
  "supportedHosts": ["SharePointWebPart"]
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
            if ((this.add && manifest.supportedHosts === undefined) ||
                (!this.add && manifest.supportedHosts !== undefined)) {
                const node = this.getAstNodeFromFile(manifest, 'supportedHosts');
                this.addOccurrence(this.resolution, manifest.path, project.path, node, occurrences);
            }
        });
        if (occurrences.length > 0) {
            this.addFindingWithOccurrences(occurrences, findings);
        }
    }
}
//# sourceMappingURL=FN011011_MAN_webpart_supportedHosts.js.map