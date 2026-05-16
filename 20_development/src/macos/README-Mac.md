# BKT Legacy for Mac

## Install

1. Quit PowerPoint.
2. Double-click `install.command`.
3. If macOS blocks the script because it was downloaded from the internet, right-click `install.command`, choose **Open**, then confirm.
4. Choose where PowerPoint should load the add-in from:

   - **1** copies `BKT-Legacy.ppam` and `Templates.pptx` to the Microsoft Office Add-Ins folder. This is recommended.
   - **2** keeps the add-in in the current folder. Keep this folder in place because `Templates.pptx` must stay next to `BKT-Legacy.ppam`.

5. Open PowerPoint.
6. Go to **Tools > PowerPoint Add-ins**, click **+**, and choose `BKT-Legacy.ppam`.

   If you chose option 1, the add-in is here:

   `~/Library/Group Containers/UBF8T346G9.Office/User Content/Add-Ins`

   If you chose option 2, the add-in is in the folder where you ran `install.command`.

   If the file picker does not show the folder, press **Command-Shift-G** and paste the path.

7. Confirm the macro security prompts.

The installer always copies:

- `BKTKeyState.scpt` to `~/Library/Application Scripts/com.microsoft.Powerpoint`.

With option 1, it also copies:

- `BKT-Legacy.ppam` and `Templates.pptx` to the PowerPoint add-in folder.

## Uninstall

Remove these files and folders:

- `~/Library/Group Containers/UBF8T346G9.Office/User Content/Add-Ins/BKT-Legacy.ppam`
- `~/Library/Group Containers/UBF8T346G9.Office/User Content/Add-Ins/Templates.pptx`
- `~/Library/Application Scripts/com.microsoft.Powerpoint/BKTKeyState.scpt`
