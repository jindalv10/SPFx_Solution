'use strict';

const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');
build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

build.initialize(gulp);

// These flags when specified with gulp tasks below affect the package.json file of a solution that is dependent on this library component.
const crntConfig = build.getConfig();
const map = require('map-stream');
const util = require("gulp-util");
const packageJsonPath = crntConfig.args["packageJsonPath"];
// The tasks below to manipulate package.json of another solution are added here because of the way library components need to built and 
// referenced. These tasks facilitate that in DevOps build pipelines.
// This task removes the reference of library component in package.json file of another solution referred by packageJsonPath variable above.
gulp.task("remove-package-reference", function () { 
  return gulp.src(packageJsonPath)
    .pipe(map(function(file, done) {
      var json = JSON.parse(file.contents.toString());
      delete json.dependencies['pd-nomination-library'];            
      file.contents = new Buffer(JSON.stringify(json));
      util.log(JSON.stringify(json));
      done(null, file);
    }))
    .pipe(gulp.dest(function (file){
      util.log(file.base);
        return file.base;
    }));
});
// This task adds reference to  library component in package.json file of library (its a separate solution).
gulp.task("add-package-reference", function () {  
  return gulp.src(packageJsonPath)
    .pipe(map(function(file, done) {
      var json = JSON.parse(file.contents.toString());
      json.dependencies['pd-nomination-library'] = "0.0.1";
      file.contents = new Buffer(JSON.stringify(json));      
      util.log(JSON.stringify(json));
      done(null, file);
    }))
    .pipe(gulp.dest(function (file){
      util.log(file.base);
        return file.base;
    }));
});
