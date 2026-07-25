'use strict';

const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);
  result.set('serve', result.get('serve-deprecated'));
  return result;
};

const gulp = require('gulp');
build.initialize(gulp);

gulp.on('stop', function (e) {
  if (e.name === '_runWatch') {
    process.stdout.write('\n\x1b[32m\x1b[1m  ██████████████████████████████████████\n  ✔  DONE — REFRESH THE WORKBENCH\n  ██████████████████████████████████████\x1b[0m\n\n');
  }
});
