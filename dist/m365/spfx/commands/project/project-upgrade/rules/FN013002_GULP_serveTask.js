import { Rule } from '../../Rule.js';
export class FN013002_GULP_serveTask extends Rule {
    constructor() {
        super();
    }
    get id() {
        return 'FN013002';
    }
    get title() {
        return 'gulpfile.js serve task';
    }
    get description() {
        return `Before 'build.initialize(require('gulp'));' add the serve task`;
    }
    get resolution() {
        return `var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};
`;
    }
    get resolutionType() {
        return 'js';
    }
    get severity() {
        return 'Required';
    }
    get file() {
        return './gulpfile.js';
    }
    visit(project, findings) {
        if (!project.gulpfileJs) {
            return;
        }
        if (project.gulpfileJs.source.indexOf(`result.set('serve', result.get('serve-deprecated'));`) < 0) {
            this.addFinding(findings);
        }
    }
}
//# sourceMappingURL=FN013002_GULP_serveTask.js.map