import { JsonRule } from '../../JsonRule.js';
export class FN003005_CFG_localizedResource_pathLib extends JsonRule {
    get id() {
        return 'FN003005';
    }
    get title() {
        return 'Update path of the localized resource';
    }
    get description() {
        return 'In the config.json file, update the path of the localized resource';
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
    get file() {
        return './config/config.json';
    }
    visit(project, findings) {
        if (!project.configJson ||
            !project.configJson.localizedResources) {
            return;
        }
        const occurrences = [];
        Object.keys(project.configJson.localizedResources).forEach(k => {
            const path = project.configJson.localizedResources[k];
            if (path.indexOf('lib/') !== 0) {
                const resolution = { localizedResources: {} };
                resolution.localizedResources[k] = `lib/${path}`;
                const node = this.getAstNodeFromFile(project.configJson, `localizedResources.${k}`);
                occurrences.push({
                    file: this.file,
                    resolution: JSON.stringify(resolution, null, 2),
                    position: this.getPositionFromNode(node)
                });
            }
        });
        if (occurrences.length > 0) {
            this.addFindingWithOccurrences(occurrences, findings);
        }
    }
}
//# sourceMappingURL=FN003005_CFG_localizedResource_pathLib.js.map