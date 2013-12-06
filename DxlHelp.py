import sublime
import sublime_plugin
import subprocess
import os.path
import winreg


BASE_PATH = os.path.abspath(os.path.dirname(__file__))


def OpenDxlHelp(text):
    helpFiles = [os.path.join(BASE_PATH, "Help\\dxl.chm")]
    try:
        doorsKey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 'SOFTWARE\\Telelogic\\DOORS', 0, winreg.KEY_READ)
        versionCount = winreg.QueryInfoKey(doorsKey)[0]
        for index in xrange(0, versionCount):
            try:
                keyName = winreg.EnumKey(doorsKey, index)
                configRegKey = winreg.OpenKey(doorsKey, keyName + '\\Config', 0, winreg.KEY_READ)
                RegValue, RegType = winreg.QueryValueEx(configRegKey, "Help System")
                helpFiles.append(os.path.join(RegValue, "dxl.chm"))
            except:
                pass
    except:
        pass

    while helpFiles:
        try:
            helpFilePath = helpFiles.pop()
            with open(helpFilePath, 'r'):
                fullPath = os.path.join(BASE_PATH, "Help\\keyhh.exe") + " -#klink " + text + " " + helpFilePath
                subprocess.Popen(fullPath)
            sublime.status_message(text)
            break
        except:
            pass


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
