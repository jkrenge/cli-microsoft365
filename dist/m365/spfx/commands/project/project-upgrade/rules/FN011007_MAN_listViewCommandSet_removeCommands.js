import { ManifestRule } from "./ManifestRule.js";
export class FN011007_MAN_listViewCommandSet_removeCommands extends ManifestRule {
    get id() {
        return 'FN011007';
    }
    get title() {
        return 'List view command set commands property';
    }
    get description() {
        return 'In the manifest remove the commands property';
    }
    get resolution() {
        return '';
    }
    get resolutionType() {
        return 'json';
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
                !commandSetManifest.commands) {
                return;
            }
            const node = this.getAstNodeFromFile(manifest, 'commands');
            this.addOccurrence(JSON.stringify({ commands: commandSetManifest.commands }, null, 2), manifest.path, project.path, node, occurrences);
        });
        if (occurrences.length > 0) {
            this.addFindingWithOccurrences(occurrences, findings);
        }
    }
}
//# sourceMappingURL=FN011007_MAN_listViewCommandSet_removeCommands.js.map