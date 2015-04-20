#!/usr/bin/env python
# -*- coding: utf-8 -*-
from __future__ import print_function

import codecs
import os
import os.path
import pytest
import stat
import subprocess
import sys

if sys.version_info.major == 2:
    from StringIO import StringIO
else:
    from io import StringIO

import off2txt.options
import off2txt.off2txt

class Cmd(object):
    def __init__(self, cmd, chdir='', indir='', stdin=None, stdout=None, stderr=None, exitcode=0):
        self.cmd = cmd
        self.chdir = chdir
        self.indir = indir
        # set and rewind stdin
        self.stdin = stdin
        if stdin:
            self.stdin = StringIO()
            for f in stdin:
                self.stdin.write(f + '\n')
            self.stdin.seek(0)
        self.stdout = stdout
        if stdout:
            if chdir:
                self.stdout = []
                for o in stdout:
                    self.stdout.append(os.path.join(chdir, o))
            elif indir:
                self.stdout = []
                for o in stdout:
                    self.stdout.append(os.path.join(indir, o))
        self.stderr = stderr
        self.exitcode = exitcode

    def __str__(self):
        return self.cmd

    def argv(self):
        args = []
        if self.chdir:
            args.extend(['-d', self.chdir])
        args.extend(self.cmd.split())
        return args

    def cmdline(self):
        args = ['off2txt']
        args.extend(self.argv())
        return args

    def run(self):
        rval = Cmd(self.cmd)
        argv = self.argv()
        stdout = StringIO()
        stderr = StringIO()
        rval.exitcode = off2txt.off2txt.main(argv, stdin=self.stdin, stdout=stdout, stderr=stderr)
        stdout_value = stdout.getvalue()
        stderr_value = stderr.getvalue()
        if stdout_value:
            rval.stdout = stdout_value.strip().split('\n')
        if stderr_value:
            rval.stderr = stderr_value.strip().split('\n')
        stdout.close()
        stderr.close()
        if self.stdin:
            self.stdin.close()
        return rval

    def run_as_process(self):
        rval = Cmd(self.cmd)
        try:
            cmd = self.cmdline()
            os.environ['COVERAGE_PROCESS_START'] = '1'
            env = os.environ.copy()
            env['COVERAGE_FILE'] = '.coverage.%s' % (self.cmd.replace('/', '-').replace(' ', '-'))
            p = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, env=env)
        except Exception as e:
            pytest.fail(msg='Cmd: run: exception running %s: %s' % (cmd, e))
        else:
            stdout, stderr = p.communicate()
            if stdout:
                rval.stdout = stdout.decode().strip().split('\n')
            if stderr:
                rval.stderr = stderr.decode().strip().split('\n')
            rval.exitcode = p.wait()
        return rval

    def good_out_file(self):
        filename = self.argv()[-1]
        base, ext = os.path.splitext(filename)
        if ext == '.xlsx':
            ext = 'csv'
        else:
            ext = 'txt'
        return base.replace('in', 'out') + '.' + ext

    def test_dir(self):
        argv = self.argv()
        i = 0
        while i < len(argv):
            if argv[i] == '-d':
                return argv[i + 1]
            i += 1
        return ''
                
    def test_out_file(self):
        base = os.path.basename(self.good_out_file())
        return os.path.join(self.test_dir(), base)
    
    
    def compare_files(self, good, test):
        with codecs.open(good, 'r', 'utf8') as good_fp:
            good_lines = good_fp.readlines()
            
        with codecs.open(test, 'r', 'utf8') as test_fp:
            test_lines = test_fp.readlines()
        
        good_length = len(good_lines)
        test_length = len(test_lines)
        if good_length != test_length:
            pytest.fail(msg='Files are different lengths: %s(%d)/%s(%d)' % (good_out_file, good_length, test_out_file, test_length))
            return False
        
        i = 0
        while i < good_length:
            if good_lines[i] != test_lines[i]:
                pytest.fail(msg='Lines are different')
                return False
            i += 1
        
        return True
    
    def split_filename(self, filename):
        base, ext = os.path.split(filename)
        ascii_filename = base + '-ascii' + ext
        unicode_filename = base + '-unicode' + ext
        return ascii_filename, unicode_filename
    
    
    def test_split_files(self, good, test):
        if os.path.exists(good):
            if not self.compare_files(good, test):
                return False
        else:
            if os.path.exists(test):
                pytest.fail(msg='File exists but should not: %s' % test)
                return False
        return True
    
    
    def compare_files_split(self, good, test):
        good_ascii, good_unicode = self.split_filename(good)
        test_ascii, test_unicode = self.split_filename(test)
        
        if not self.test_split_files(good_ascii, test_ascii):
            return False
        
        if not self.test_split_files(good_unicode, test_unicode):
            return False
    
        return True
    
    
    def is_split(self):
        for a in self.argv():
            if a == '-s':
                return True
        return False
    
    
    def files_match(self):
        good_out_file = self.good_out_file()
        test_out_file = self.test_out_file()
        if self.is_split():
            return self.compare_files_split(good_out_file, test_out_file)
        return self.compare_files(good_out_file, test_out_file)


some_bad_option = '--some-bad-option'
usage_string_expanded = 'usage: %s' % (off2txt.options.usage_string % {'prog': off2txt.options.program_name})
some_bad_option_error_msg = '%(prog)s: error: unrecognized arguments: %(option)s' % {'prog': off2txt.options.program_name, 'option': some_bad_option}


cmds = [
    # word
    Cmd('-X -d tests-out/docx tests/docx/in/01.docx'),
    Cmd('-X -d tests-out/docx -s tests/docx/in/01.docx'),
    Cmd('-X -d tests-out/docx tests/docx/in/02.docx'),
    Cmd('-X -d tests-out/docx -s tests/docx/in/02.docx'),
    Cmd('-X -d tests-out/docx tests/docx/in/03.docx'),
    Cmd('-X -d tests-out/docx -s tests/docx/in/03.docx'),
    
    # PowerPoint
    Cmd('-X -d tests-out/pptx tests/pptx/in/01.pptx'),
    Cmd('-X -d tests-out/pptx -s tests/pptx/in/01.pptx'),
    Cmd('-X -d tests-out/pptx tests/pptx/in/02.pptx'),
    Cmd('-X -d tests-out/pptx -s tests/pptx/in/02.pptx'),
    Cmd('-X -d tests-out/pptx tests/pptx/in/03.pptx'),
    Cmd('-X -d tests-out/pptx -s tests/pptx/in/03.pptx'),
    
    # Excel
    Cmd('-X -d tests-out/xlsx tests/xlsx/in/01.xlsx'),
    Cmd('-X -d tests-out/xlsx -s tests/xlsx/in/01.xlsx'),
    Cmd('-X -d tests-out/xlsx tests/xlsx/in/02.xlsx'),
    Cmd('-X -d tests-out/xlsx -s tests/xlsx/in/02.xlsx'),
    Cmd('-X -d tests-out/xlsx tests/xlsx/in/03.xlsx'),
    Cmd('-X -d tests-out/xlsx -s tests/xlsx/in/03.xlsx'),
]

cmds_as_process = [

]


class TestCmd(object):
    @pytest.mark.parametrize('cmd', cmds)
    def test_cmd(self, cmd):
        r = cmd.run()
        assert r.stderr == cmd.stderr
        assert r.stdout == cmd.stdout
        assert r.exitcode == cmd.exitcode
        assert cmd.files_match()

    @pytest.mark.parametrize('cmd', cmds_as_process)
    def test_cmd_as_process(self, cmd):
        r = cmd.run_as_process()
        assert r.stderr == cmd.stderr
        assert r.stdout == cmd.stdout
        assert r.exitcode == cmd.exitcode
        assert cmd.files_match()


if __name__ == '__main__':
    pytest.main()


