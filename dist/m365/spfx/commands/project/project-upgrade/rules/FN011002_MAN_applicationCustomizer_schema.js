import { ManifestRule } from "./ManifestRule.js";
export class FN011002_MAN_applicationCustomizer_schema extends ManifestRule {
    constructor(schema) {
        super();
        this.schema = schema;
    }
    get id() {
        return 'FN011002';
    }
    get title() {
        return 'Application customizer manifest schema';
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
            if (manifest.componentType === 'Extension' &&
                manifest.extensionType === 'ApplicationCustomizer' &&
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
//# sourceMappingURL=FN011002_MAN_applicationCustomizer_schema.js.map