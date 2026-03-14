import { ManifestRule } from "./ManifestRule.js";
export class FN011010_MAN_webpart_version extends ManifestRule {
    get id() {
        return 'FN011010';
    }
    get title() {
        return 'Web part manifest version';
    }
    get description() {
        return 'Update version in manifest to use automated component versioning';
    }
    get resolution() {
        return `{
  "version": "*",
}`;
    }
    get severity() {
        return 'Optional';
    }
    visit(project, findings) {
        if (!project.manifests ||
            project.manifests.length === 0) {
            return;
        }
        const occurrences = [];
        project.manifests.forEach(manifest => {
            if (manifest.componentType === 'WebPart' &&
                manifest.version !== '*') {
                const node = this.getAstNodeFromFile(manifest, 'version');
                this.addOccurrence(this.resolution, manifest.path, project.path, node, occurrences);
            }
        });
        if (occurrences.length > 0) {
            this.addFindingWithOccurrences(occurrences, findings);
        }
    }
}
//# sourceMappingURL=FN011010_MAN_webpart_version.js.map