import { ManifestRule } from "./ManifestRule.js";
export class FN011001_MAN_webpart_schema extends ManifestRule {
    constructor(schema) {
        super();
        this.schema = schema;
    }
    get id() {
        return 'FN011001';
    }
    get title() {
        return 'Web part manifest schema';
    }
    get description() {
        return 'Update schema in manifest';
    }
    get resolution() {
        return `{
  "$schema": "${this.schema}"
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
            if (manifest.componentType === 'WebPart' &&
                manifest.$schema !== this.schema) {
                const node = this.getAstNodeFromFile(manifest, '$schema');
                this.addOccurrence(this.resolution, manifest.path, project.path, node, occurrences);
            }
        });
        if (occurrences.length > 0) {
            this.addFindingWithOccurrences(occurrences, findings);
        }
    }
}
//# sourceMappingURL=FN011001_MAN_webpart_schema.js.map