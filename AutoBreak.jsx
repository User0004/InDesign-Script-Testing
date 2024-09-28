// NAME:AutoBreak
// STATUS: Ready 
// FUNCTION: Splits a selected text frame with multiple columns into separate linked columns in Adobe InDesign, while preserving the text flow and layout. 
// Note: The script will not work on sensitive first box database items for header, body, and captions.
// Prefered keyboard shortcut is alt+F1
// Version 1.0.0

// AutoBreak
function start() {
    var selectedFrame = app.activeDocument.selection[0];

    // Check if a selection was made and if it is a TextFrame
    if (selectedFrame == undefined || selectedFrame.constructor.name != 'TextFrame') {
        alert("Please select a text frame.");
        return;
    }

    // Check if the frame is the first in the thread
    if (selectedFrame.previousTextFrame == null) {
        alert("Cannot process the first text frame.");
        return;
    }

    // Check if the selected frame has at least 2 columns
    if (selectedFrame.textFramePreferences.textColumnCount < 2) {
        alert("Please select a text frame that has two or more columns.");
        return;
    }

    // Call the createColumnsFromSelection function with the selected frame
    createColumnsFromSelection(selectedFrame);
}

// Execute the start function as a script with specific settings
app.doScript(start, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, 'Organize Columns');

// Define the function to create columns from the selected frame
function createColumnsFromSelection(frame) {
    // Set the measurement units to points for consistency
    app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;

    // Initialize variables for columns, gutter, bounds, width, insets, and frame list
    var columns = frame.textFramePreferences.textColumnCount,
        gutter = frame.textFramePreferences.textColumnGutter,
        bounds = frame.geometricBounds,
        width = bounds[3] - bounds[1],
        insets = frame.textFramePreferences.insetSpacing,
        frameList = [frame];

    // Calculate the width of each column
    var colWidth = (width - insets[1] - insets[3] - (gutter * (columns - 1))) / columns;

    // Configure the original frame to become the first column
    frame.textFramePreferences.textColumnCount = 1;
    frame.textFramePreferences.insetSpacing = [insets[0], 0, insets[2], 0];
    frame.geometricBounds = [bounds[0], bounds[1] + insets[1], bounds[2], bounds[1] + insets[1] + colWidth];

    // Loop to create additional columns
    for (var i = 1; i < columns; i++) {
        // Duplicate the previous frame and adjust its position
        var newFrame = frameList[i - 1].duplicate(undefined, [(colWidth + gutter), 0]);

        // Clear the content of the new frame
        newFrame.parentStory.contents = '';

        // Link the new frame to the previous frame using nextTextFrame
        frameList[i - 1].nextTextFrame = newFrame;

        // Add the new frame to the frame list for tracking
        frameList.push(newFrame);
    }
}
