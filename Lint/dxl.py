import re

from base_linter import BaseLinter, os, INPUT_METHOD_TEMP_FILE

CONFIG = {
    'language': 'DXL',
    'lint_args': "{filename}",
    'input_method': INPUT_METHOD_TEMP_FILE
}


class Linter(BaseLinter):
    ERROR_RE = re.compile(r'^-E- DXL: <(.*):(?P<line>[0-9]+)> (?P<error>.*)$')

    def get_executable(self, view):
        return (True, os.path.join(self.LIB_PATH, 'dxl', "DxlLint.exe"), "")

    def parse_errors(self, view, errors, lines, errorUnderlines,
                     violationUnderlines, warningUnderlines,
                     errorMessages, violationMessages,
                     warningMessages):
        # Go through each line in the output of checkDXL
        for line in errors.splitlines():
            match = self.ERROR_RE.match(line)
            if match:
                line, error = match.group('line'), match.group('error')
                lineno = int(line)
                self.add_message(lineno, lines, error, errorMessages)
