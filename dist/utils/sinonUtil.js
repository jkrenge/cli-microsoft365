export const sinonUtil = {
    restore(method) {
        if (!method) {
            return;
        }
        if (!Array.isArray(method)) {
            method = [method];
        }
        method.forEach((m) => {
            if (m && m.restore) {
                m.restore();
            }
        });
    }
};
//# sourceMappingURL=sinonUtil.js.map