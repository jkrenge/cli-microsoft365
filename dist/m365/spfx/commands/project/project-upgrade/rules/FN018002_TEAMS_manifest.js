import fs from 'fs';
import path from 'path';
import { Rule } from '../../Rule.js';
export class FN018002_TEAMS_manifest extends Rule {
    constructor() {
        super();
    }
    get id() {
        return 'FN018002';
    }
    get title() {
        return 'Web part Microsoft Teams tab manifest';
    }
    get description() {
        return 'Create Microsoft Teams tab manifest for the web part';
    }
    get resolution() {
        return `add_cmd[BEFOREPATH]"__filePath__"[AFTERPATH][BEFORECONTENT]
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.2/MicrosoftTeams.schema.json",
  "manifestVersion": "1.2",
  "packageName": "__packageName__",
  "id": "__id__",
  "version": "0.1",
  "developer": {
    "name": "SPFx + Teams Dev",
    "websiteUrl": "https://products.office.com/en-us/sharepoint/collaboration",
    "privacyUrl": "https://privacy.microsoft.com/en-us/privacystatement",
    "termsOfUseUrl": "https://www.microsoft.com/en-us/servicesagreement"
  },
  "name": {
    "short": "__name__"
  },
  "description": {
    "short": "__description__",
    "full": "__description__"
  },
  "icons": {
    "outline": "tab20x20.png",
    "color": "tab96x96.png"
  },
  "accentColor": "#004578",
  "configurableTabs": [
    {
      "configurationUrl": "https://{teamSiteDomain}{teamSitePath}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest={teamSitePath}/_layouts/15/teamshostedapp.aspx%3FopenPropertyPane=true%26teams%26componentId=__id__",
      "canUpdateConfiguration": true,
      "scopes": [
        "team"
      ]
    }
  ],
  "validDomains": [
    "*.login.microsoftonline.com",
    "*.sharepoint.com",
    "*.sharepoint-df.com",
    "spoppe-a.akamaihd.net",
    "spoprod-a.akamaihd.net",
    "resourceseng.blob.core.windows.net",
    "msft.spoppe.com"
  ],
  "webApplicationInfo": {
    "resource": "https://{teamSiteDomain}",
    "id": "00000003-0000-0ff1-ce00-000000000000"
  }
}
[AFTERCONTENT]`;
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
            const webPartFolderName = path.basename(path.dirname(manifest.path));
            const teamsFolderName = `teams`;
            const teamsFolderPath = path.join(project.path, teamsFolderName);
            const teamsManifestPath = path.join(teamsFolderPath, `manifest_${webPartFolderName}.json`);
            if (fs.existsSync(teamsManifestPath)) {
                return;
            }
            let webPartTitle = 'undefined';
            let webPartDescription = 'undefined';
            if (manifest.preconfiguredEntries &&
                manifest.preconfiguredEntries.length > 0) {
                const entry = manifest.preconfiguredEntries[0];
                if (entry.title && entry.title.default) {
                    webPartTitle = entry.title.default;
                }
                if (entry.description && entry.description.default) {
                    webPartDescription = entry.description.default;
                }
            }
            const webPartId = manifest.id || 'undefined';
            const resolution = this.resolution
                .replace(/__filePath__/g, teamsManifestPath)
                .replace(/__packageName__/g, webPartTitle)
                .replace(/__id__/g, webPartId)
                .replace(/__name__/g, webPartTitle)
                .replace(/__description__/g, webPartDescription);
            occurrences.push({
                file: path.relative(project.path, teamsManifestPath),
                resolution: resolution
            });
        });
        if (occurrences.length > 0) {
            this.addFindingWithOccurrences(occurrences, findings);
        }
    }
}
//# sourceMappingURL=FN018002_TEAMS_manifest.js.map