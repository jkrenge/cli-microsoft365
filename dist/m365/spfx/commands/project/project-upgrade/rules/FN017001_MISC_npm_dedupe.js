import { Rule } from '../../Rule.js';
export class FN017001_MISC_npm_dedupe extends Rule {
    constructor() {
        super();
    }
    get id() {
        return 'FN017001';
    }
    get title() {
        return 'Run npm dedupe';
    }
    get description() {
        return `If, after upgrading npm packages, when building the project you have errors similar to: "error TS2345: Argument of type 'SPHttpClientConfiguration' is not assignable to parameter of type 'SPHttpClientConfiguration'", try running 'npm dedupe' to cleanup npm packages.`;
    }
    get resolution() {
        return 'npm dedupe';
    }
    get resolutionType() {
        return 'cmd';
    }
    get file() {
        return './package.json';
    }
    get severity() {
        return 'Optional';
    }
    visit(project, findings) {
        const npmFinding = findings.find(f => typeof f.occurrences.find(o => o.resolution.indexOf('install') === 0 || o.resolution.indexOf('uninstall') === 0) !== 'undefined');
        if (npmFinding) {
            this.addFinding(findings);
        }
    }
}
//# sourceMappingURL=FN017001_MISC_npm_dedupe.js.map