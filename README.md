# DXL package for Sublime Text

This package adds support for DOORS eXtension Language (DXL) to [Sublime Text] [ST2] for:
* Syntax Highlighting
* Snippets
* Build Configuration

## Installation

There are several ways to install this package.

### Package Control

The easiest way to install this package is through [Package Control] [PC].

* Install [Package Control] [PC]
* Open the `Command Palette` (`Tools >> Command Palette` or `Ctrl+Shift+P` or `Cmd+Shift+P`).
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

### DXL Lint

Plugin for [SublimeLinter] [LINT] to add support for DXL.

* Install [SublimeLinter] [LINT]
* Copy `\DXL\Lint\dxl.py` to `\SublimeLinter\sublimelinter\modules\dxl.py`
* Create `\SublimeLinter\sublimelinter\modules\libs\dxl\`
* Copy `\DXL\Lint\DxlLint.exe` to `\SublimeLinter\sublimelinter\modules\libs\dxl\DxlLint.exe`

### Syntax Highlighting Colour Schemes

Modified version of the Soda Dark version of Monokai to add support for the DXL Language and Lint.

## Credits
Original author: Adam Cadamally

## Licence

	Copyright (c) 2013 Adam Cadamally

	Permission is hereby granted, free of charge, to any person obtaining a copy
	of this software and associated documentation files (the "Software"), to deal
	in the Software without restriction, including without limitation the rights
	to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
	copies of the Software, and to permit persons to whom the Software is
	furnished to do so, subject to the following conditions:

	The above copyright notice and this permission notice shall be included in
	all copies or substantial portions of the Software.

	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
	IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
	FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
	AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
	LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
	THE SOFTWARE.

 [ST2]: http://www.sublimetext.com/
 [PC]: http://wbond.net/sublime_packages/package_control
 [LINT]: https://github.com/SublimeLinter/SublimeLinter
