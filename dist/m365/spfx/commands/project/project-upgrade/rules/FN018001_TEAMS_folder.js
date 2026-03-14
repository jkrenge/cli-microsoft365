import fs from 'fs';
import path from 'path';
import { Rule } from '../../Rule.js';
export class FN018001_TEAMS_folder extends Rule {
    constructor() {
        super();
    }
    get id() {
        return 'FN018001';
    }
    get title() {
        return 'Web part Microsoft Teams tab resources folder';
    }
    get description() {
        return 'Create folder for Microsoft Teams tab resources';
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
        const teamsFolderName = 'teams';
        const teamsFolderPath = path.join(project.path, teamsFolderName);
        if (!fs.existsSync(teamsFolderPath)) {
            this.addFindingWithCustomInfo(this.title, this.description, [{
                    file: path.relative(project.path, teamsFolderPath),
                    resolution: `create_dir_cmdPathParam${project.path}NameParam${teamsFolderName}ItemTypeParam`
                }], findings);
        }
    }
}
//# sourceMappingURL=FN018001_TEAMS_folder.js.map