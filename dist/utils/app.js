import { createRequire } from 'module';
const require = createRequire(import.meta.url);
let packageJson;
export const app = {
    packageJson: () => {
        if (!packageJson) {
            packageJson = require('../../package.json');
        }
        return packageJson;
    }
};
//# sourceMappingURL=app.js.map