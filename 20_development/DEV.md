# BKT Legacy Toolbar Development

## Master files for development

All code and ribbon settings for the addin must be located in the file `20_development/BKT-Legacy.pptm`. For proper tracking of changes, all code is also exported into the following files:

* VBA code: `20_development/src/*`
* Ribbon XML: `20_development/ribbonUI.xml`

PPTM file and exported files should remain in sync.


## VBA Import/Export

There is a separate addin to import and export all VBA code in/from the PPTM file. The addin is located in `20_development/VBAImportExport/VBAImportExport.ppam`. When installed it will add two buttons to the developer tab.

The Export button will copy all VBA modules of the currently opened file into a `src` folder next to the open file. The Import button will delete all existing VBA modules and create new modules based on the `src` folder next to the open file.

## Editing VBA code

New functionalities can be developed within the PPTM file using the VBA code editor. Afterwards the code should be exported to track changes. All changes should be committed in both PPTM file and export files at the same time. In case of deviations, the exported files are the master.

## Editing ribbon settings

Any changes to the ribbon (custom UI configuration) need to be done in the `ribbonUI.xml` file. Aftwards, the script `ribbonUI-import-from-xml` should be executed (required a unix shell) to import the XML file into the PPTM file.

Alternatively, the open source [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor) can be used to edit the XML file and also remove or add image files. The ribbonUI.xml file should also be in sync with the PPTM file.

## Compile addin

For distribution the addin must be compiled into a PPAM file. Open the PPTM file, select Save as and choose "PowerPoint-Add_in (\*.ppam)" as file type. New releases are saved in `30_builds/10_stable`. The previous version should be moved to `30_builds/99_archive`.
