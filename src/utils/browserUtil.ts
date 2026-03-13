import { networkAccess } from './networkAccess.js';

export const browserUtil = {
  /* c8 ignore next 5 */
  async open(url: string): Promise<void> {
    if (url.startsWith('http://') || url.startsWith('https://')) {
      networkAccess.assertAllowed(url);
    }

    const _open = (await import('open')).default;
    const runningOnWindows = process.platform === 'win32';
    await _open(url, { wait: runningOnWindows });
  }
};
