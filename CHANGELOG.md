# Changelog

## [1.7.0] - 2026-05-15

### New Features

- Added a library function for quickly opening PowerPoint files and wired the template feature to the new library helper.
- Added pickup/apply reference bounds so position or size can be copied from one shape to others.
- Added support for swapping more than two shapes at once.
- Added new same width and same height modes based on average and last selected shape.
- Added process shape creation.
- Added replace-keep-size when replacing shapes.
- Added typography helpers for protected hyphens, protected spaces, and protected narrow spaces.
- Added cleanup for unused designs.
- Added stretch-by-last arrange actions.
- Added centered sticker placement.
- Added an option to set all text margins equally.
- Added support for using the slide master content box as arrange/stretch reference when only one shape is selected.

### Improvements

- Improved Mac modifier key support with AppleScript integration, optional key handling toggle, install script, and reduced permission prompts.
- Improved spinner behavior with larger Shift-based steps and better reset behavior for separation controls using Ctrl/Cmd.
- Improved arrange, stretch, and same width/height actions to respect rotated shapes more reliably.
- Improved language setting so it applies to the current selection or selected slides instead of all slides by default.
- Improved selection-by-shape tools so they work with multiple selected shapes.
- Improved performance by limiting AppleScript key checks to controls that actually need modifier key state.
- Improved Mac key-state performance when several modifier checks are needed at once.
- Improved internal ribbon and action handling through broader refactoring, shared selection helpers, and reduced shape-by-shape loops where shape range operations can be used directly.
- Refreshed and expanded icon assets for the newer commands and updated the UI to hide font box labels and move version display into code.

### Bug Fixes

- Fixed icon registration in the import script.
- Fixed the enabled state for saving selected slides.
- Fixed Mac save-selected-slides support.
- Fixed stretch-to-slide-master behavior.
- Fixed the fixed point behavior not working correctly.
