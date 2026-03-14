import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraGroupSettingTemplateListCommand extends GraphCommand {
    get name() {
        return commands.GROUPSETTINGTEMPLATE_LIST;
    }
    get description() {
        return 'Lists Entra group settings templates';
    }
    defaultProperties() {
        return ['id', 'displayName'];
    }
    async commandAction(logger) {
        try {
            const templates = await odata.getAllItems(`${this.resource}/v1.0/groupSettingTemplates`);
            await logger.log(templates);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new EntraGroupSettingTemplateListCommand();
//# sourceMappingURL=groupsettingtemplate-list.js.map