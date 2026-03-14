import { FileTokenStorage } from './FileTokenStorage.js';
class MsalCachePlugin {
    constructor() {
        this.fileTokenStorage = new FileTokenStorage(FileTokenStorage.msalCacheFilePath());
    }
    async beforeCacheAccess(tokenCacheContext) {
        try {
            const data = await this.fileTokenStorage.get();
            tokenCacheContext.tokenCache.deserialize(data);
        }
        catch {
            // Do nothing
        }
    }
    async afterCacheAccess(tokenCacheContext) {
        if (!tokenCacheContext.cacheHasChanged) {
            return;
        }
        try {
            await this.fileTokenStorage.set(tokenCacheContext.tokenCache.serialize());
        }
        catch {
            // Do nothing
        }
    }
}
const msalCachePlugin = new MsalCachePlugin();
export { msalCachePlugin };
//# sourceMappingURL=msalCachePlugin.js.map