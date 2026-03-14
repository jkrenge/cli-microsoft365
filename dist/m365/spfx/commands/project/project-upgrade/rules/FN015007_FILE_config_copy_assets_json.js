import { FileAddRemoveRule } from "./FileAddRemoveRule.js";
export class FN015007_FILE_config_copy_assets_json extends FileAddRemoveRule {
    constructor(add) {
        super('./config/copy-assets.json', add);
    }
    get id() {
        return 'FN015007';
    }
    get supersedes() {
        return ['FN004001', 'FN004002'];
    }
}
//# sourceMappingURL=FN015007_FILE_config_copy_assets_json.js.map