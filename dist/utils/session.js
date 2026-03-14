import * as crypto from 'crypto';
import { cache } from './cache.js';
export const session = {
    getId(pid) {
        const key = `${pid.toString()}_session`;
        let sessionId = cache.getValue(key);
        if (sessionId) {
            return sessionId;
        }
        sessionId = crypto.randomBytes(24).toString('base64');
        cache.setValue(key, sessionId);
        return sessionId;
    }
};
//# sourceMappingURL=session.js.map