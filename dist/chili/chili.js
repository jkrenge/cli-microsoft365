import fs from 'fs';
import path from 'path';
import url from 'url';
import request from '../request.js';
import { md } from '../utils/md.js';
import { prompt } from '../utils/prompt.js';
const __dirname = url.fileURLToPath(new URL('.', import.meta.url));
const mendableBaseUrl = 'https://api.mendable.ai/v1';
const mendableApiKey = 'd3313d54-6f8e-40e0-90d3-4095019d4be7';
let showHelp = false;
let debug = false;
let conversationId = 0;
let initialPrompt = '';
let history = [];
request.logger = {
    /* c8 ignore next 3 */
    log: async (msg) => console.log(msg),
    logRaw: async (msg) => console.log(msg),
    logToStderr: async (msg) => console.error(msg)
};
request.debug = debug;
function getPromptFromArgs(args) {
    showHelp = args.indexOf('--help') > -1 || args.indexOf('-h') > -1;
    if (showHelp) {
        const commandsFolder = path.join(__dirname, '..', 'm365');
        const pathChunks = [commandsFolder, '..', '..', 'docs', 'docs', 'user-guide', 'chili.mdx'];
        const helpFilePath = path.join(...pathChunks);
        if (fs.existsSync(helpFilePath)) {
            let helpContents = fs.readFileSync(helpFilePath, 'utf8');
            helpContents = md.md2plain(helpContents, path.join(commandsFolder, '..', '..', 'docs'));
            console.log(helpContents);
            return '';
        }
        else {
            console.error('Help file not found');
            return '';
        }
    }
    else {
        // reset to default. needed for tests
        showHelp = false;
    }
    const debugPos = args.indexOf('--debug');
    if (debugPos > -1) {
        debug = true;
        request.debug = true;
        args.splice(debugPos, 1);
    }
    else {
        // reset to default. needed for tests
        debug = false;
    }
    return args.join(' ');
}
async function startConversation(args) {
    history = [];
    initialPrompt = getPromptFromArgs(args);
    if (showHelp) {
        return;
    }
    conversationId = await getConversationId();
    if (!initialPrompt) {
        initialPrompt = await promptForPrompt();
    }
    await runConversationTurn(conversationId, initialPrompt);
}
async function promptForPrompt() {
    return await prompt.forInput({ message: '🌶️  How can I help you?' });
}
async function runConversationTurn(conversationId, question) {
    console.log('');
    const response = await runMendableChat(conversationId, question);
    history.push({
        prompt: question,
        response: response.answer.text
    });
    console.log(md.md2plain(response.answer.text, ''));
    console.log('');
    console.log('Source:');
    // remove duplicates
    const sources = response.sources.filter((src, index, self) => index === self.findIndex(s => s.link === src.link));
    sources.forEach(src => console.log(`⬥ ${src.link}`));
    console.log('');
    const choices = [
        {
            name: '📝 I want to know more',
            value: 'ask'
        },
        {
            name: '👋 I know enough. Thanks!',
            value: 'end'
        },
        {
            name: '🔄 I want to ask about something else',
            value: 'new'
        }
    ];
    const result = await prompt.forSelection({ message: 'What would you like to do next?', choices });
    switch (result) {
        case 'ask': {
            const prompt = await promptForPrompt();
            await runConversationTurn(conversationId, prompt);
            break;
        }
        case 'end':
            await endConversation(conversationId);
            console.log('');
            console.log('🌶️   Bye!');
            break;
        case 'new':
            initialPrompt = '';
            await startConversation([]);
            break;
    }
}
async function endConversation(conversationId) {
    const requestOptions = {
        url: `${mendableBaseUrl}/endConversation`,
        headers: {
            'content-type': 'application/json',
            'x-anonymous': true
        },
        responseType: 'json',
        data: {
            // eslint-disable-next-line camelcase
            api_key: mendableApiKey,
            // eslint-disable-next-line camelcase
            conversation_id: conversationId
        }
    };
    await request.post(requestOptions);
}
async function runMendableChat(conversationId, question) {
    const requestOptions = {
        url: `${mendableBaseUrl}/mendableChat`,
        headers: {
            'content-type': 'application/json',
            'x-anonymous': true
        },
        responseType: 'json',
        data: {
            // eslint-disable-next-line camelcase
            api_key: mendableApiKey,
            // eslint-disable-next-line camelcase
            conversation_id: conversationId,
            question,
            history,
            shouldStream: false
        }
    };
    return await request.post(requestOptions);
}
async function getConversationId() {
    const requestOptions = {
        url: `${mendableBaseUrl}/newConversation`,
        headers: {
            'content-type': 'application/json',
            'x-anonymous': true
        },
        responseType: 'json',
        data: {
            // eslint-disable-next-line camelcase
            api_key: mendableApiKey
        }
    };
    const response = await request.post(requestOptions);
    return response.conversation_id;
}
export const chili = {
    startConversation
};
//# sourceMappingURL=chili.js.map