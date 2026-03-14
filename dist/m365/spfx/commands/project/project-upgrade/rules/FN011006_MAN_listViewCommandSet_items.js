import { ManifestRule } from "./ManifestRule.js";
export class FN011006_MAN_listViewCommandSet_items extends ManifestRule {
    get id() {
        return 'FN011006';
    }
    get title() {
        return 'List view command set items property';
    }
    get description() {
        return 'In the manifest add the items property';
    }
    get resolution() {
        return '';
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
            const commandSetManifest = manifest;
            if (commandSetManifest.componentType !== 'Extension' ||
                commandSetManifest.extensionType !== 'ListViewCommandSet' ||
                commandSetManifest.items) {
                return;
            }
            const resolution = {
                items: commandSetManifest.commands || {}
            };
            Object.keys(resolution.items).forEach(k => {
                resolution.items[k].title = { default: resolution.items[k].title };
                resolution.items[k].type = 'command';
            });
            const node = this.getAstNodeFromFile(manifest, 'commands');
            this.addOccurrence(JSON.stringify(resolution, null, 2), manifest.path, project.path, node, occurrences);
        });
        if (occurrences.length > 0) {
            this.addFindingWithOccurrences(occurrences, findings);
        }
    }
}
//# sourceMappingURL=FN011006_MAN_listViewCommandSet_items.js.map