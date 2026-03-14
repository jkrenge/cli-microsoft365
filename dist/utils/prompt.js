import { settingsNames } from '../settingsNames.js';
import { cli } from '../cli/cli.js';
let inquirerInput;
let inquirerConfirm;
let inquirerSelect;
;
;
export const prompt = {
    /* c8 ignore next 16 */
    async forInput(config) {
        if (!inquirerInput) {
            inquirerInput = await import('@inquirer/input');
        }
        const errorOutput = cli.getSettingWithDefaultValue(settingsNames.errorOutput, 'stderr');
        return inquirerInput
            .default(config, { output: errorOutput === 'stderr' ? process.stderr : process.stdout })
            .catch(error => {
            if (error instanceof Error && error.name === 'ExitPromptError') {
                return ''; // noop; handle Ctrl + C
            }
            throw error;
        });
    },
    /* c8 ignore next 9 */
    async forConfirmation(config) {
        if (!inquirerConfirm) {
            inquirerConfirm = await import('@inquirer/confirm');
        }
        const errorOutput = cli.getSettingWithDefaultValue(settingsNames.errorOutput, 'stderr');
        return inquirerConfirm.default(config, { output: errorOutput === 'stderr' ? process.stderr : process.stdout });
    },
    /* c8 ignore next 14 */
    async forSelection(config) {
        if (!inquirerSelect) {
            inquirerSelect = await import('@inquirer/select');
        }
        const errorOutput = cli.getSettingWithDefaultValue(settingsNames.errorOutput, 'stderr');
        const promptPageSizeCap = cli.getSettingWithDefaultValue(settingsNames.promptListPageSize, 7);
        if (!config.pageSize) {
            config.pageSize = Math.min(config.choices.length, promptPageSizeCap);
        }
        return inquirerSelect.default(config, { output: errorOutput === 'stderr' ? process.stderr : process.stdout });
    }
};
//# sourceMappingURL=prompt.js.map