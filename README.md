# DXL package for Sublime Text

This package adds support for DOORS eXtension Language (DXL) to [Sublime Text] [ST2] for:

* Syntax Highlighting
* Snippets
* Build Configuration
* Jump to DXL Keyword in Help

## Installation

There are several ways to install this package.

### Package Control

The easiest way to install this package is through [Package Control] [PC].

* Install [Package Control] [PC]
* Open the `Command Palette` (`Tools >> Command Palette`).
* Type `Install Package` and hit return.
* Type `DXL` and hit return.

### Using Git

Go to your Sublime Text `Packages` directory and clone the repository using the command below:

    git clone https://github.com/SublimeText/DXL "DXL"

### Download Manually

* Download the files using the GitHub .zip download option
* Unzip the files and rename the folder to `DXL`
* Copy the folder to your Sublime Text `Packages` directory

## Bonus Features

### DXL Help

Use `F1` to jump to the current word in the `dxl.chm` file.

### DXL Lint

Plugin for [SublimeLinter] [LINT] to add support for DXL.

* Install [SublimeLinter] [LINT]
* Copy `\DXL\Lint\dxl.py` to `\SublimeLinter\sublimelinter\modules\dxl.py`
* Create `\SublimeLinter\sublimelinter\modules\libs\dxl\`
* Copy `\DXL\Lint\DxlLint.exe` to `\SublimeLinter\sublimelinter\modules\libs\dxl\DxlLint.exe`

### Syntax Highlighting Colour Schemes

Modified version of the Soda Dark version of Monokai to add support for the DXL Language and Lint.

## Unicode

If you have problems with unicode characters in the Build Output, set Microsoft Windows to use UTF8 on the command line:

* Run `regedit`
* Navigate to Key `HKEY_LOCAL_MACHINE\Software\Microsoft\Command Processor\`
* Set Value name `Autorun` to Value data `chcp 65001 > nul`

## Say Thanks

Donate: [GitTip] [TIP] 

 [ST2]: http://www.sublimetext.com/
 [PC]: http://wbond.net/sublime_packages/package_control
 [LINT]: https://github.com/SublimeLinter/SublimeLinter
 [TIP]: https://www.gittip.com/Adarma/
