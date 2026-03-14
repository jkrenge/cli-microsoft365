import { spfx } from '../../../../../../utils/spfx.js';
import { FileAddRemoveRule } from "./FileAddRemoveRule.js";
export class FN015008_FILE_eslintrc_js extends FileAddRemoveRule {
    constructor(add, contents) {
        super('./.eslintrc.js', add, contents);
    }
    get id() {
        return 'FN015008';
    }
    visit(project, notifications) {
        if (spfx.isReactProject(project)) {
            this.contents = this.contents.replace('@microsoft/eslint-config-spfx/lib/profiles/default', '@microsoft/eslint-config-spfx/lib/profiles/react');
        }
        super.visit(project, notifications);
    }
}
//# sourceMappingURL=FN015008_FILE_eslintrc_js.js.map