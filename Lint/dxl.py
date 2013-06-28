from base_linter import BaseLinter, os, re, INPUT_METHOD_TEMP_FILE

CONFIG = {
    'language': 'DXL',
    'lint_args': "{filename}",
    'input_method': INPUT_METHOD_TEMP_FILE
}


class Linter(BaseLinter):
    ERROR_RE = re.compile(r'^-E- DXL: <(?P<path>.*):(?P<line>[0-9]+)> (?P<error>.*)$')
    STACK_RE = re.compile(r'\s*<(?P<path>.*):(?P<line>[0-9]+)>\s*$')

    def get_executable(self, view):
        return (True, os.path.join(self.LIB_PATH, 'dxl', "DxlLint.exe"), "")

    def parse_errors(self, view, errors, lines, errorUnderlines,
                     violationUnderlines, warningUnderlines,
                     errorMessages, violationMessages,
                     warningMessages):
        bufferName = "\\.tempfiles\\" + os.path.basename(view.file_name())
        error = None
        # Go through each line in the output of checkDXL
        for fullline in errors.splitlines():
            match = self.ERROR_RE.match(fullline)
            if match:
                path, line, error = match.group('path'), match.group('line'), match.group('error')
                if path.endswith(bufferName):
                    lineno = int(line)
                    self.add_message(lineno, lines, error, errorMessages)
                    error = None  # don't report this error again
            if error:  # if error is defined from earlier when you matched ERROR_RE
                match = self.STACK_RE.match(fullline)
                if match:
                    callpath, callline = match.group('path'), match.group('line')
                    if callpath.endswith(bufferName):
                        calllineno = int(callline)
                        prefix = "<" + os.path.basename(path) + ":" + line + "> "
                        self.add_message(calllineno, lines, prefix + error, errorMessages)
                        error = None  # don't report this error again
