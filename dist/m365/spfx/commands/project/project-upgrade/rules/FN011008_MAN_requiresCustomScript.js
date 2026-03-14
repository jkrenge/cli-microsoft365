import { ManifestRule } from "./ManifestRule.js";
export class FN011008_MAN_requiresCustomScript extends ManifestRule {
    get id() {
        return 'FN011008';
    }
    get supersedes() {
        return ['FN011009'];
    }
    get title() {
        return 'Client-side component manifest requiresCustomScript property';
    }
    get description() {
        return 'In the manifest rename the safeWithCustomScriptDisabled property to requiresCustomScript and invert its value';
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
            if (typeof manifest.safeWithCustomScriptDisabled !== 'undefined') {
                const node = this.getAstNodeFromFile(manifest, 'safeWithCustomScriptDisabled');
                this.addOccurrence(`{
  "requiresCustomScript": ${!manifest.safeWithCustomScriptDisabled}
}`, manifest.path, project.path, node, occurrences);
            }
        });
        if (occurrences.length > 0) {
            this.addFindingWithOccurrences(occurrences, findings);
        }
    }
}
//# sourceMappingURL=FN011008_MAN_requiresCustomScript.js.map