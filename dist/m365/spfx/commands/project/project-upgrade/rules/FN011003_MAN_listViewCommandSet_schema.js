import { ManifestRule } from "./ManifestRule.js";
export class FN011003_MAN_listViewCommandSet_schema extends ManifestRule {
    constructor(schema) {
        super();
        this.schema = schema;
    }
    get id() {
        return 'FN011003';
    }
    get title() {
        return 'List view command set manifest schema';
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
                manifest.extensionType === 'ListViewCommandSet' &&
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
//# sourceMappingURL=FN011003_MAN_listViewCommandSet_schema.js.map