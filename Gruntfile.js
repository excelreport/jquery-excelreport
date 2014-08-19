module.exports = function(grunt) {
    'use strict';
    function loadDependencies(deps) {
        if (deps) {
            for (var key in deps) {
                if (key.indexOf("grunt-") == 0) {
                    grunt.loadNpmTasks(key);
                }
            }
        }
    }
 
    grunt.initConfig({
        pkg: grunt.file.readJSON('package.json'),

        clean: {
            dist: "dist"
        },
        copy: {
            dist: {
                files: [{
                    expand: true,
                    cwd: "src/i18n",
                    src: "*.js",
                    dest: "dist/i18n"
                }]
            },
            app: {
                files: [{
                    expand: true,
                    cwd: "dist",
                    src: "*.*",
                    dest: "../report2/public/javascripts/client"
                },
                {
                    expand: true,
                    cwd: "dist/i18n",
                    src: "*.js",
                    dest: "../report2/public/javascripts/client/i18n"
                }]
            }
        },
        concat: {
            core : {
                src : [
                    "src/intro.txt",
                    "src/constant.js",
                    "src/enum.js",
                    "src/functions.js",
                    "src/ruleManager.js",
                    "src/popup.js",
                    "src/processor.js",
                    "src/excelReport.js",
                    "src/jqueryImpl.js",
                    "src/outro.txt"
                ],
                dest: "target/jquery.excelreport.js"
            },
            css : {
                src : [
                    "bower_components/excel2canvas/dist/jquery.excel2canvas.css",
                    "src/jquery.excelreport.css"
                ],
                dest: "dist/jquery.excelreport.full.css"
            },
            full : {
                src : [
                    "bower_components/flotr2/flotr2.js",
                    "bower_components/roomframework/dist/roomframework.js",
                    "bower_components/excel2canvas/dist/jquery.excel2canvas.js",
                    "bower_components/excel2canvas/dist/jquery.excel2chart.flotr2.js",
                    "target/jquery.excelreport.js"
                ],
                dest: "dist/jquery.excelreport.full.js"
            },
            nochart : {
                src : [
                    "bower_components/roomframework/dist/roomframework.js",
                    "bower_components/excel2canvas/dist/jquery.excel2canvas.js",
                    "target/jquery.excelreport.js"
                ],
                dest: "dist/jquery.excelreport.nochart.js"
            }
        },

        uglify: {
            build: {
                files: [{
                    "dist/jquery.excelreport.full.min.js": "dist/jquery.excelreport.full.js",
                    "dist/jquery.excelreport.nochart.min.js": "dist/jquery.excelreport.nochart.js",
                    "dist/i18n/jquery.excelreport.msg_ja.min.js": "dist/i18n/jquery.excelreport.msg_ja.js"
                }]
            }
        },
        cssmin: {
            minify: {
                expand: true,
                cwd: 'dist',
                src: ['*.css', '!*.min.css'],
                dest: 'dist',
                ext: '.min.css',
                extDot: 'last'
            }
        },

        jshint : {
            all : ['src/*.js']
        },
        
        watch: {
            scripts: {
                files: [
                    'src/*.js',
                    'src/*.css'
                ],
                tasks: ['jshint', 'concat:core', 'concat:css', 'copy:dist', 'concat:full', 'concat:nochart', 'uglify', 'cssmin', "copy:app"]
            }
        }
    });
 
    loadDependencies(grunt.config("pkg").devDependencies);

    grunt.registerTask('default', [ 'jshint', 'concat:core', 'concat:css', 'copy:dist', 'concat:full', 'concat:nochart', 'uglify', 'cssmin']);
    grunt.registerTask('cp', [ 'copy:app']);

};