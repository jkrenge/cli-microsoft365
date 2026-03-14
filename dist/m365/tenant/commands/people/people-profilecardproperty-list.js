import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { odata } from '../../../../utils/odata.js';
class TenantPeopleProfileCardPropertyListCommand extends GraphCommand {
    get name() {
        return commands.PEOPLE_PROFILECARDPROPERTY_LIST;
    }
    get description() {
        return 'Lists all profile card properties';
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Listing all profile card properties...`);
            }
            const result = await odata.getAllItems(`${this.resource}/v1.0/admin/people/profileCardProperties`);
            let output = result;
            if (args.options.output && args.options.output !== 'json') {
                output = output.sort((n1, n2) => {
                    const localizations1 = n1.annotations[0]?.localizations?.length ?? 0;
                    const localizations2 = n2.annotations[0]?.localizations?.length ?? 0;
                    if (localizations1 > localizations2) {
                        return -1;
                    }
                    if (localizations1 < localizations2) {
                        return 1;
                    }
                    return 0;
                });
                output = output.map((p) => {
                    const propertyAnnotations = p.annotations[0]?.localizations?.map((l) => {
                        return { ['displayName ' + l.languageTag]: l.displayName };
                    }) ?? [];
                    const propertyAnnotationsObject = Object.assign({}, ...propertyAnnotations);
                    const result = { directoryPropertyName: p.directoryPropertyName };
                    if (p.annotations[0]?.displayName) {
                        result.displayName = p.annotations[0]?.displayName;
                    }
                    return {
                        ...result,
                        ...propertyAnnotationsObject
                    };
                });
            }
            await logger.log(output);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new TenantPeopleProfileCardPropertyListCommand();
//# sourceMappingURL=people-profilecardproperty-list.js.map