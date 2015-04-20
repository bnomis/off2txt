# -*- coding: utf-8 -*-
from __future__ import print_function

import argparse
import sys

from . import __version__


# put these string here so we can import them for testing
program_name = 'off2txt'
usage_string = '%(prog)s [options] File [File ...]'
version_string = '%(prog)s %(version)s' % {'prog': program_name, 'version': __version__}
description_string = '''off2txt: extract ASCII/Unicode text from Office files to separate files
'''


def parse_opts(argv, stdin=None, stdout=None, stderr=None):
    parser = argparse.ArgumentParser(
        prog=program_name,
        usage=usage_string,
        description=description_string,
        formatter_class=argparse.RawDescriptionHelpFormatter
    )

    parser.add_argument(
        '--version',
        action='version',
        version=version_string
    )

    parser.add_argument(
        '--debug',
        dest='debug',
        action='store_true',
        default=False,
        help='Turn on debug logging.'
    )

    parser.add_argument(
        '--debug-log',
        dest='debug_log',
        metavar='FILE',
        help='Save debug logging to FILE.'
    )

    parser.add_argument(
        '-a',
        '--ascii',
        default='ascii',
        metavar='EXTENSION',
        help='Identifier to append to input file name to make ASCII output file name when splitting Unicode and ASCII text. Default %(default)s.'
    )

    parser.add_argument(
        '-d',
        '--directory',
        metavar='DIRECTORY',
        help='Save extracted text to DIRECTORY. Ignored if the -o option is given.'
    )

    parser.add_argument(
        '-e',
        '--extension',
        default='txt',
        help='Extension to use for extracted text files. Default for Word and PowerPoint is %(default)s. Default for Excel is csv.'
    )

    parser.add_argument(
        '-o',
        '--output',
        metavar='FILE',
        help='Save extracted text to FILE. If not given, the output file is named the same as the input file but with a txt extension. The extension can be changed with the -e option. Files are opened in append mode unless the -X option is given.'
    )

    parser.add_argument(
        '-s',
        '--split',
        default=False,
        action='store_true',
        help='Split ASCII and Unicode text into two separate files. Unicode files are named by adding -unicode before the file extension. The Unicode identifer can be changed with the -u option.'
    )

    parser.add_argument(
        '-u',
        '--unicode',
        default='unicode',
        metavar='EXTENSION',
        help='Identifier to append to input file name to make Unicode output file name when splitting Unicode and ASCII text. Default %(default)s.'
    )

    parser.add_argument(
        '-A',
        '--suppress-file-access-errors',
        default=False,
        action='store_true',
        help='Do not print file/directory access errors.'
    )

    parser.add_argument(
        '-X',
        '--overwrite-output-files',
        default=False,
        action='store_true',
        help='Truncate output files before writing.'
    )

    parser.add_argument(
        'files',
        metavar='File',
        nargs='+',
        help='Files to extract from'
    )

    # print('argv = %s' % argv)
    options = parser.parse_args(argv)

    # set up i/o options
    options.stdin = stdin or sys.stdin
    options.stdout = stdout or sys.stdout
    options.stderr = stderr or sys.stderr

    options.did_extract = False
    options.exit_status = 'not-set'

    return options

