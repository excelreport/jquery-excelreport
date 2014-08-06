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

        copy: {
            dist: {
                files: [{
                    expand: true,
                    cwd: "src",
                    src: "*.js",
                    dest: "dist"
                }]
            },
            app: {
                files: [{
                    expand: true,
                    cwd: "dist",
                    src: "*.js",
                    dest: "../report2/public/javascripts"
                }]
            }
        },

        uglify: {
            build: {
                files: [{
                    "dist/jquery.excelreport.min.js": "dist/jquery.excelreport.js",
                    "dist/jquery.excelreport.msg_ja.min.js": "dist/jquery.excelreport.msg_ja.js",
                }]
            }
        },

        jshint : {
            all : ['src/*.js']
        },
        
        watch: {
            scripts: {
                files: [
                    'src/*.js'
                ],
                tasks: ['jshint', 'copy:dist', 'uglify', "copy:app"]
            }
        }
    });
 
    loadDependencies(grunt.config("pkg").devDependencies);

    grunt.registerTask('default', [ 'jshint', 'copy:dist', 'uglify']);
    grunt.registerTask('cp', [ 'copy:app']);

};