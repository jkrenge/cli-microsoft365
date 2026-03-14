import { Rule } from '../../Rule.js';
import { FN018002_TEAMS_manifest } from './FN018002_TEAMS_manifest.js';
export class FN018005_TEAMS_deprecated_manifest extends Rule {
    constructor() {
        super();
    }
    get id() {
        return 'FN018005';
    }
    get supersedes() {
        return ['FN018002'];
    }
    get title() {
        // title is kept empty so that the finding isn't reported
        return '';
    }
    get description() {
        return `Manually creating Microsoft Teams manifests for web parts is no longer necessary because they're created automatically`;
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
        // this rule should be applied whenever the FN018002_TEAMS_manifest is
        const deprecatedRule = new FN018002_TEAMS_manifest();
        const fn018002Findings = [];
        deprecatedRule.visit(project, fn018002Findings);
        if (fn018002Findings.length > 0) {
            this.addFinding(findings);
        }
    }
}
//# sourceMappingURL=FN018005_TEAMS_deprecated_manifest.js.map