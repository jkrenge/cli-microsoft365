import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    chatId: z.string()
        .refine(id => validation.isValidTeamsChatId(id), {
        error: e => `'${e.input}' is not a valid value for option chatId.`
    })
        .alias('i'),
    endDateTime: z.string()
        .refine(time => validation.isValidISODateTime(time), {
        error: e => `'${e.input}' is not a valid ISO date-time string for option endDateTime.`
    })
        .optional()
});
class TeamsChatMessageListCommand extends GraphCommand {
    get name() {
        return commands.CHAT_MESSAGE_LIST;
    }
    get description() {
        return 'Lists all messages from a chat';
    }
    defaultProperties() {
        return ['id', 'createdDateTime', 'shortBody'];
    }
    get schema() {
        return options;
    }
    async commandAction(logger, args) {
        try {
            let apiUrl = `${this.resource}/v1.0/chats/${args.options.chatId}/messages`;
            if (args.options.endDateTime) {
                // You can only filter results if the request URL contains the $orderby and $filter query parameters configured for the same property;
                // otherwise, the $filter query option is ignored.
                apiUrl += `?$filter=createdDateTime lt ${args.options.endDateTime}&$orderby=createdDateTime desc`;
            }
            const items = await odata.getAllItems(apiUrl);
            if (args.options.output && args.options.output !== 'json') {
                items.forEach(i => {
                    // hoist the content to body for readability
                    i.body = i.body.content;
                    let shortBody;
                    const bodyToProcess = i.body;
                    if (bodyToProcess) {
                        let maxLength = 50;
                        let addedDots = '...';
                        if (bodyToProcess.length < maxLength) {
                            maxLength = bodyToProcess.length;
                            addedDots = '';
                        }
                        shortBody = bodyToProcess.replace(/\n/g, ' ').substring(0, maxLength) + addedDots;
                    }
                    i.shortBody = shortBody;
                });
            }
            await logger.log(items);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new TeamsChatMessageListCommand();
//# sourceMappingURL=chat-message-list.js.map