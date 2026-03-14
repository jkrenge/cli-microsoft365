import { JsonRule } from '../../JsonRule.js';
export class FN003003_CFG_bundles extends JsonRule {
    get id() {
        return 'FN003003';
    }
    get title() {
        return `config.json bundles`;
    }
    get description() {
        return `In config.json add the 'bundles' property`;
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
            typeof project.configJson.bundles !== 'undefined') {
            return;
        }
        const entries = project.configJson.entries;
        if (!entries) {
            return;
        }
        const resolution = { bundles: {} };
        entries.forEach((e) => {
            resolution.bundles[this.tryGetBundleName(e.entry)] = {
                components: [{
                        entrypoint: e.entry,
                        manifest: e.manifest
                    }]
            };
        });
        const node = this.getAstNodeFromFile(project.configJson, 'bundles');
        this.addFindingWithOccurrences([{
                file: this.file,
                resolution: JSON.stringify(resolution, null, 2),
                position: this.getPositionFromNode(node)
            }], findings);
    }
    /**
     * Smart guess on the bundle name.
     * @param entry the entry value
     */
    tryGetBundleName(entry) {
        let name = entry.substring(entry.lastIndexOf('/') + 1, entry.length);
        name = name.replace('.js', '');
        name = name.replace(/([a-z](?=[A-Z]))/g, '$1-');
        name = name.toLowerCase();
        return name;
    }
}
//# sourceMappingURL=FN003003_CFG_bundles.js.map