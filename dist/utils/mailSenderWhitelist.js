import fs from 'fs';
import os from 'os';
import path from 'path';
import { cli } from '../cli/cli.js';
import { settingsNames } from '../settingsNames.js';
function getFilePath() {
    return path.join(os.homedir(), '.cli-m365-mail-sender-whitelist.json');
}
function getEmptyStore() {
    return {
        domains: [],
        senders: []
    };
}
function normalizeValue(value) {
    const normalizedValue = value?.trim().toLowerCase();
    return normalizedValue ? normalizedValue : undefined;
}
function normalizeStore(store) {
    const domains = Array.isArray(store?.domains)
        ? Array.from(new Set(store.domains.map(domain => normalizeValue(domain)).filter((domain) => typeof domain !== 'undefined'))).sort()
        : [];
    const senders = Array.isArray(store?.senders)
        ? Array.from(new Set(store.senders.map(sender => normalizeValue(sender)).filter((sender) => typeof sender !== 'undefined'))).sort()
        : [];
    return {
        domains,
        senders
    };
}
function load() {
    const filePath = getFilePath();
    if (!fs.existsSync(filePath)) {
        return getEmptyStore();
    }
    try {
        return normalizeStore(JSON.parse(fs.readFileSync(filePath, 'utf8')));
    }
    catch {
        return getEmptyStore();
    }
}
function save(store) {
    fs.writeFileSync(getFilePath(), JSON.stringify(normalizeStore(store), null, 2), 'utf8');
}
function getSenderAddress(message) {
    return normalizeValue(message.from?.emailAddress?.address) ??
        normalizeValue(message.sender?.emailAddress?.address);
}
function getSenderDomain(message) {
    const senderAddress = getSenderAddress(message);
    if (!senderAddress) {
        return undefined;
    }
    const addressParts = senderAddress.split('@');
    return addressParts.length === 2 ? addressParts[1] : undefined;
}
function isAllowedSender(message, store = load()) {
    const senderAddress = getSenderAddress(message);
    if (!senderAddress) {
        return false;
    }
    if (store.senders.includes(senderAddress)) {
        return true;
    }
    const senderDomain = getSenderDomain(message);
    return typeof senderDomain !== 'undefined' && store.domains.includes(senderDomain);
}
function filterMessages(messages) {
    const store = load();
    const allowed = [];
    const blocked = [];
    messages.forEach(message => {
        if (isAllowedSender(message, store)) {
            allowed.push(message);
            return;
        }
        blocked.push(message);
    });
    return {
        allowed,
        blocked,
        filteredCount: blocked.length
    };
}
function looksLikeMailMessage(value) {
    return !!value &&
        typeof value === 'object' &&
        (typeof value.from?.emailAddress?.address === 'string' ||
            typeof value.sender?.emailAddress?.address === 'string') &&
        (typeof value.subject === 'string' ||
            typeof value.receivedDateTime === 'string');
}
function filterResponse(response) {
    if (Array.isArray(response) && response.some(looksLikeMailMessage)) {
        const store = load();
        let filteredCount = 0;
        const filteredResponse = response.filter(item => {
            if (!looksLikeMailMessage(item)) {
                return true;
            }
            if (isAllowedSender(item, store)) {
                return true;
            }
            filteredCount++;
            return false;
        });
        return {
            filteredResponse: filteredResponse,
            filteredCount
        };
    }
    if (response &&
        typeof response === 'object' &&
        Array.isArray(response.value) &&
        response.value.some(looksLikeMailMessage)) {
        const value = response.value;
        const result = filterMessages(value);
        const filteredResponse = {
            ...response,
            value: result.allowed
        };
        if (result.allowed.length === 0) {
            delete filteredResponse['@odata.nextLink'];
        }
        return {
            filteredResponse,
            filteredCount: result.filteredCount
        };
    }
    if (looksLikeMailMessage(response)) {
        const isAllowed = isAllowedSender(response);
        return {
            filteredResponse: (isAllowed ? response : {}),
            filteredCount: isAllowed ? 0 : 1
        };
    }
    return {
        filteredResponse: response,
        filteredCount: 0
    };
}
function shouldFilterInInteractiveMode() {
    return cli.getSettingWithDefaultValue(settingsNames.prompt, true) &&
        !!process.stdout.isTTY &&
        !!process.stdin.isTTY;
}
async function logFilteredSummary(logger, filteredCount) {
    if (filteredCount === 0) {
        return;
    }
    await logger.logToStderr(`Filtered ${filteredCount} message(s) from non-whitelisted senders. Use 'm365 outlook message whitelist' to review them.`);
}
function addSender(senderAddress) {
    const normalizedSenderAddress = normalizeValue(senderAddress);
    if (!normalizedSenderAddress) {
        return false;
    }
    const store = load();
    if (store.senders.includes(normalizedSenderAddress)) {
        return false;
    }
    store.senders.push(normalizedSenderAddress);
    save(store);
    return true;
}
function addDomain(domain) {
    const normalizedDomain = normalizeValue(domain);
    if (!normalizedDomain) {
        return false;
    }
    const store = load();
    if (store.domains.includes(normalizedDomain)) {
        return false;
    }
    store.domains.push(normalizedDomain);
    save(store);
    return true;
}
export const mailSenderWhitelist = {
    addDomain,
    addSender,
    filterMessages,
    filterResponse,
    getFilePath,
    getSenderAddress,
    getSenderDomain,
    isAllowedSender,
    load,
    logFilteredSummary,
    shouldFilterInInteractiveMode
};
//# sourceMappingURL=mailSenderWhitelist.js.map