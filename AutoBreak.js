// NAME:Autobreak

// Version ------------- NOTE ------------- This script will work on the first text frames part of thread -------- the updated version prevents this so consider that script instead of this for Naviga

// Status: Working  - conditional - do not use on first Naviga box as you cannot duplicate database items - for other instances works
// FUNCTION - To break text frames spanning =>2 columns into the their respective single columns while maintaing threading
// OUTLINE - reads the size of columns, duplicates columns and threads 
// Prefered keyboard shortcut is F5

// Define the start function to initiate the column organization
function start() {
    // Call the createColumnsFromSelection function with the first item from the selection
    createColumnsFromSelection(app.activeDocument.selection[0]);
}

// Execute the start function as a script with specific settings
app.doScript(start, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, 'Organize Columns');

// Define the function to create columns from the selected frame
function createColumnsFromSelection(frame) {
    // Check if the provided frame is valid for columnization
    if (
        frame == undefined ||
        frame.constructor.name != 'TextFrame' ||
        frame.textFramePreferences.textColumnCount < 2
    )
        return; // Exit the function if the frame is invalid or doesn't meet column count criteria

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

