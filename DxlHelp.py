import sublime
import sublime_plugin
import subprocess
import os.path

BASE_PATH = os.path.abspath(os.path.dirname(__file__))


def OpenDxlHelp(text):
    full_path = os.path.join(BASE_PATH, "Help\\keyhh.exe") + " -#klink " + text + " " + os.path.join(BASE_PATH, "Help\\dxl.chm")
    subprocess.Popen(full_path)
    sublime.status_message(text)


class DxlKeywordHelpCommand(sublime_plugin.TextCommand):
    def run(self, edit):
        curr_view = self.view
        curr_sel = curr_view.sel()[0]
        if curr_view.match_selector(curr_sel.begin(), 'source.dxl'):

            word_end = curr_sel.end()
            if curr_sel.empty():
                word = curr_view.substr(curr_view.word(word_end)).lower()
            else:
                word = curr_view.substr(curr_sel).lower()
            if word is None or len(word) <= 1:
                sublime.status_message('No function selected')

            OpenDxlHelp(word)
        else:
            sublime.status_message("No Help Available")
