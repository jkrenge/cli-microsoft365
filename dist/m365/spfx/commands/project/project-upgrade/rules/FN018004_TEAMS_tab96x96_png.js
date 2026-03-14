import fs from 'fs';
import path from 'path';
import url from 'url';
import { Rule } from '../../Rule.js';
const __dirname = url.fileURLToPath(new URL('.', import.meta.url));
export class FN018004_TEAMS_tab96x96_png extends Rule {
    /**
     * Creates instance of this rule
     * @param fixedFileName Name to use for the copied file. If not specified, will generate the name based on web part's ID
     */
    constructor(fixedFileName) {
        super();
        this.fixedFileName = fixedFileName;
    }
    get id() {
        return 'FN018004';
    }
    get title() {
        return 'Web part Microsoft Teams tab large icon';
    }
    get description() {
        return 'Create Microsoft Teams tab large icon for the web part';
    }
    get resolution() {
        return '';
    }
    get resolutionType() {
        return 'cmd';
    }
    get file() {
        return '';
    }
    get severity() {
        return 'Optional';
    }
    visit(project, findings) {
        if (!project.manifests ||
            project.manifests.length < 1) {
            return;
        }
        const webPartManifests = project.manifests.filter(m => m.componentType === 'WebPart');
        if (webPartManifests.length < 1) {
            return;
        }
        const occurrences = [];
        webPartManifests.forEach(manifest => {
            if (!manifest.id) {
                return;
            }
            const teamsFolderName = `teams`;
            const teamsFolderPath = path.join(project.path, teamsFolderName);
            const iconName = this.getIconName(manifest);
            const iconPath = path.join(teamsFolderPath, iconName);
            if (!fs.existsSync(iconPath)) {
                occurrences.push({
                    file: path.relative(project.path, iconPath),
                    resolution: `copy_cmd "${path.join(__dirname, '..', 'assets', 'tab96x96.png')}"DestinationParam"${iconPath}"`
                });
            }
        });
        if (occurrences.length > 0) {
            this.addFindingWithOccurrences(occurrences, findings);
        }
    }
    getIconName(manifest) {
        if (this.fixedFileName) {
            return this.fixedFileName;
        }
        return `${manifest.id}_color.png`;
    }
}
//# sourceMappingURL=FN018004_TEAMS_tab96x96_png.js.map