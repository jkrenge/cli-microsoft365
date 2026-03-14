import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
export const chatUtil = {
    /**
     * Finds existing Microsoft Teams chats by participants, using the Microsoft Graph
     * @param expectedMemberEmails a string array of participant email addresses
     * @param logger a logger to pipe into the graph request odata helper.
     */
    async findExistingChatsByParticipants(expectedMemberEmails) {
        const chatType = expectedMemberEmails.length === 2 ? 'oneOnOne' : 'group';
        const endpoint = `https://graph.microsoft.com/v1.0/chats?$filter=chatType eq '${chatType}'&$expand=members&$select=id,topic,createdDateTime,members`;
        const foundChats = [];
        const chats = await odata.getAllItems(endpoint);
        for (const chat of chats) {
            const chatMembers = chat.members;
            if (chatMembers.length === expectedMemberEmails.length) {
                const chatMemberEmails = chatMembers.map(member => member.email?.toLowerCase());
                if (expectedMemberEmails.every(email => chatMemberEmails.some(memberEmail => memberEmail === email))) {
                    foundChats.push(chat);
                }
            }
        }
        return foundChats;
    },
    /**
     * Finds existing Microsoft Teams chats by name, using the Microsoft Graph
     * @param name the name of the chat conversation to find
     * @param logger a logger to pipe into the graph request odata helper.
     */
    async findExistingGroupChatsByName(name) {
        const endpoint = `https://graph.microsoft.com/v1.0/chats?$filter=topic eq '${formatting.encodeQueryParameter(name).replace("'", "''")}'&$expand=members&$select=id,topic,createdDateTime,chatType`;
        return odata.getAllItems(endpoint);
    }
};
//# sourceMappingURL=chatUtil.js.map