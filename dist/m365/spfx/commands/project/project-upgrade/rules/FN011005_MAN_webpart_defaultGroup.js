import { ManifestRule } from "./ManifestRule.js";
export class FN011005_MAN_webpart_defaultGroup extends ManifestRule {
    constructor(oldDefaultGroup, newDefaultGroup) {
        super();
        this.oldDefaultGroup = oldDefaultGroup;
        this.newDefaultGroup = newDefaultGroup;
    }
    get id() {
        return 'FN011005';
    }
    get title() {
        return 'Web part manifest default group';
    }
    get description() {
        return 'In the manifest update the default group value';
    }
    get resolution() {
        return `{
  "preconfiguredEntries": [{
    "group": { "default": "${this.newDefaultGroup}" }
  }]
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
                manifest.preconfiguredEntries) {
                manifest.preconfiguredEntries.forEach((e, i) => {
                    if (e.group && e.group.default === this.oldDefaultGroup) {
                        const node = this.getAstNodeFromFile(manifest, `preconfiguredEntries[${i}].group.default`);
                        this.addOccurrence(this.resolution, manifest.path, project.path, node, occurrences);
                    }
                });
            }
        });
        if (occurrences.length > 0) {
            this.addFindingWithOccurrences(occurrences, findings);
        }
    }
}
//# sourceMappingURL=FN011005_MAN_webpart_defaultGroup.js.map