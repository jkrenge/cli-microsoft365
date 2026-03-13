import fs from 'fs';
import os from 'os';
import path from 'path';
import { Logger } from '../cli/Logger.js';
import { cli } from '../cli/cli.js';
import { settingsNames } from '../settingsNames.js';

type MessageParticipant = {
  emailAddress?: {
    address?: string | null;
  } | null;
} | null;

export interface MailMessageLike {
  id?: string;
  subject?: string | null;
  receivedDateTime?: string | null;
  sender?: MessageParticipant | null;
  from?: MessageParticipant | null;
}

interface MailSenderWhitelistStore {
  domains: string[];
  senders: string[];
}

interface MailFilterResult<T> {
  allowed: T[];
  blocked: T[];
  filteredCount: number;
}

function getFilePath(): string {
  return path.join(os.homedir(), '.cli-m365-mail-sender-whitelist.json');
}

function getEmptyStore(): MailSenderWhitelistStore {
  return {
    domains: [],
    senders: []
  };
}

function normalizeValue(value?: string | null): string | undefined {
  const normalizedValue = value?.trim().toLowerCase();
  return normalizedValue ? normalizedValue : undefined;
}

function normalizeStore(store: Partial<MailSenderWhitelistStore> | undefined): MailSenderWhitelistStore {
  const domains = Array.isArray(store?.domains)
    ? Array.from(new Set(store!.domains.map(domain => normalizeValue(domain)).filter((domain): domain is string => typeof domain !== 'undefined'))).sort()
    : [];
  const senders = Array.isArray(store?.senders)
    ? Array.from(new Set(store!.senders.map(sender => normalizeValue(sender)).filter((sender): sender is string => typeof sender !== 'undefined'))).sort()
    : [];

  return {
    domains,
    senders
  };
}

function load(): MailSenderWhitelistStore {
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

function save(store: MailSenderWhitelistStore): void {
  fs.writeFileSync(getFilePath(), JSON.stringify(normalizeStore(store), null, 2), 'utf8');
}

function getSenderAddress(message: MailMessageLike): string | undefined {
  return normalizeValue(message.from?.emailAddress?.address) ??
    normalizeValue(message.sender?.emailAddress?.address);
}

function getSenderDomain(message: MailMessageLike): string | undefined {
  const senderAddress = getSenderAddress(message);
  if (!senderAddress) {
    return undefined;
  }

  const addressParts = senderAddress.split('@');
  return addressParts.length === 2 ? addressParts[1] : undefined;
}

function isAllowedSender(message: MailMessageLike, store: MailSenderWhitelistStore = load()): boolean {
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

function filterMessages<T extends MailMessageLike>(messages: T[]): MailFilterResult<T> {
  const store = load();
  const allowed: T[] = [];
  const blocked: T[] = [];

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

function looksLikeMailMessage(value: any): value is MailMessageLike {
  return !!value &&
    typeof value === 'object' &&
    (
      typeof value.from?.emailAddress?.address === 'string' ||
      typeof value.sender?.emailAddress?.address === 'string'
    ) &&
    (
      typeof value.subject === 'string' ||
      typeof value.receivedDateTime === 'string'
    );
}

function filterResponse<T>(response: T): { filteredResponse: T; filteredCount: number; } {
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
      filteredResponse: filteredResponse as T,
      filteredCount
    };
  }

  if (response &&
    typeof response === 'object' &&
    Array.isArray((response as unknown as { value?: unknown[]; }).value) &&
    (response as unknown as { value: unknown[]; }).value.some(looksLikeMailMessage)) {
    const value = (response as unknown as { value: MailMessageLike[]; }).value;
    const result = filterMessages(value);
    const filteredResponse: any = {
      ...(response as any),
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
      filteredResponse: (isAllowed ? response : {}) as T,
      filteredCount: isAllowed ? 0 : 1
    };
  }

  return {
    filteredResponse: response,
    filteredCount: 0
  };
}

function shouldFilterInInteractiveMode(): boolean {
  return cli.getSettingWithDefaultValue<boolean>(settingsNames.prompt, true) &&
    !!process.stdout.isTTY &&
    !!process.stdin.isTTY;
}

async function logFilteredSummary(logger: Logger, filteredCount: number): Promise<void> {
  if (filteredCount === 0) {
    return;
  }

  await logger.logToStderr(`Filtered ${filteredCount} message(s) from non-whitelisted senders. Use 'm365 outlook message whitelist' to review them.`);
}

function addSender(senderAddress: string): boolean {
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

function addDomain(domain: string): boolean {
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
