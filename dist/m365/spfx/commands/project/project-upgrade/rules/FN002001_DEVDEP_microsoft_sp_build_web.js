import { DependencyRule } from "./DependencyRule.js";
export class FN002001_DEVDEP_microsoft_sp_build_web extends DependencyRule {
    constructor(packageVersion, add = true) {
        super('@microsoft/sp-build-web', packageVersion, true, false, add);
    }
    get id() {
        return 'FN002001';
    }
}
//# sourceMappingURL=FN002001_DEVDEP_microsoft_sp_build_web.js.map