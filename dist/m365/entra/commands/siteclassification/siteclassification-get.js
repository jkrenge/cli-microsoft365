import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { SiteClassificationSettings } from './SiteClassificationSettings.js';
class EntraSiteClassificationGetCommand extends GraphCommand {
    get name() {
        return commands.SITECLASSIFICATION_GET;
    }
    get description() {
        return 'Gets site classification configuration';
    }
    async commandAction(logger) {
        try {
            const requestOptions = {
                url: `${this.resource}/v1.0/groupSettings`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            const res = await request.get(requestOptions);
            if (res.value.length === 0) {
                throw 'Site classification is not enabled.';
            }
            const unifiedGroupSetting = res.value.filter((directorySetting) => {
                return directorySetting.displayName === 'Group.Unified';
            });
            if (unifiedGroupSetting === null || unifiedGroupSetting.length === 0) {
                throw "Missing DirectorySettingTemplate for \"Group.Unified\"";
            }
            const siteClassificationsSettings = new SiteClassificationSettings();
            // Get the classification list
            const classificationList = unifiedGroupSetting[0].values.filter((directorySetting) => {
                return directorySetting.name === 'ClassificationList';
            });
            siteClassificationsSettings.Classifications = [];
            if (classificationList !== null && classificationList.length > 0) {
                siteClassificationsSettings.Classifications = classificationList[0].value.split(',');
            }
            // Get the UsageGuidelinesUrl
            const guidanceUrl = unifiedGroupSetting[0].values.filter((directorySetting) => {
                return directorySetting.name === 'UsageGuidelinesUrl';
            });
            siteClassificationsSettings.UsageGuidelinesUrl = "";
            if (guidanceUrl !== null && guidanceUrl.length > 0) {
                siteClassificationsSettings.UsageGuidelinesUrl = guidanceUrl[0].value;
            }
            // Get the GuestUsageGuidelinesUrl
            const guestGuidanceUrl = unifiedGroupSetting[0].values.filter((directorySetting) => {
                return directorySetting.name === 'GuestUsageGuidelinesUrl';
            });
            siteClassificationsSettings.GuestUsageGuidelinesUrl = "";
            if (guestGuidanceUrl !== null && guestGuidanceUrl.length > 0) {
                siteClassificationsSettings.GuestUsageGuidelinesUrl = guestGuidanceUrl[0].value;
            }
            // Get the DefaultClassification
            const defaultClassification = unifiedGroupSetting[0].values.filter((directorySetting) => {
                return directorySetting.name === 'DefaultClassification';
            });
            siteClassificationsSettings.DefaultClassification = "";
            if (defaultClassification !== null && defaultClassification.length > 0) {
                siteClassificationsSettings.DefaultClassification = defaultClassification[0].value;
            }
            await logger.log(JSON.parse(JSON.stringify(siteClassificationsSettings)));
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new EntraSiteClassificationGetCommand();
//# sourceMappingURL=siteclassification-get.js.map