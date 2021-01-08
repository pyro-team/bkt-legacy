# BKT Legacy Toolbar

<img src="00_documentation/screenshot-legacy.png">

## Introduction

The BKT Legacy Toolbox is a VBA-based toolbar for Microsoft PowerPoint. It is the predecessor of the [BKT](https://github.com/pyro-team/bkt-toolbox/). Compared to the newer BKT written in Python the Legacy Toolbar has less features, various restrictions and is not extensible. However, unlike the BKT it will mostly run on Mac Office.

The BKT is developed by us in our spare time, so we cannot offer support or respond to special requests.

### Language

Historically, the BKT Legacy Toolbar was developed in German and unfortunately, we do not have the time to translate the whole toolbox. We hope that most functions are self-explanatory. If you have experience in multi-language VBA projects, feel free to support us.

## System requirements

The BKT Legacy Toolbar runs under Windows from Office 2010 in all current Office versions as well as on Mac starting with Office 2016. Some functions are not available on Mac though.

## Installation

The PowerPoint-Addin can be downloaded as [compiled `BKT-Legacy.ppam` file](https://github.com/pyro-team/bkt-legacy/releases/latest). (Same version as in `30_builds\10_stable\`.)

* On Windows it can be installed at File > Options > Add-Ins > Select "PowerPoint-Add-Ins" in Menu below > In new dialog click Add and select downloaded file.
* On Mac go to Extras > PowerPoint-Add-Ins and click "+" to select the downloaded file. Confirm any security questions to activate macros.
* In order to add template slides via menu in PowerPoint, you need to download the `Templates.pptx` file and place it next to the addin file. Feel free to add your own template slides.

## Development

To install the development version please refer to instructions on [20_development/DEV.md](20_development/DEV.md).

## Contributions

 * [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor)