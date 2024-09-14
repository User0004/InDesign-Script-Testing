// NAME: Top Shrink
// STATUS: Ready
// FUNCTION: Decrease height of selected items from top one baseline for each time the script is run
// Best to use on non group grouped items for Nav. If needed use white arrow to handle groups, but black arrow is find for text boxes 
// Prefered keyboard shortcut is Top shift F9 



//Top Shrink

app.doScript(function() {
    // Get the active document
    var doc = app.activeDocument;

    // Get document grid preferences
    var gridPreferences = doc.gridPreferences;
    var baselineIncrement = gridPreferences.baselineDivision;

    // Add a precision threshold to handle floating-point precision issues
    var precisionThreshold = 0.1; // This helps to avoid errors when height is close to baselineIncrement

    // Function to adjust height
    function adjustHeight(item) {
        var currentHeight = item.geometricBounds[2] - item.geometricBounds[0];

        if (currentHeight <= baselineIncrement + precisionThreshold) {
            // If the height is less than or equal to the baseline increment (with tolerance), move the item downwards
            var moveAmount = Math.abs(baselineIncrement); // Absolute value of baseline increment
            
            // Use the move method to move the item downwards
            item.move(undefined, [0, moveAmount]);
        } else if (currentHeight > baselineIncrement + precisionThreshold) {
            // If the height is greater than the baseline increment, reduce height from the bottom
            if (item instanceof Image && item.parent instanceof Rectangle) {
                // Adjust the height of the parent frame, not the image itself
                adjustHeight(item.parent);
            } else {
                var newHeight = currentHeight - baselineIncrement;

                // Check to ensure the new height does not go below one baseline increment
                if (newHeight < baselineIncrement) {
                    newHeight = baselineIncrement;
                }
                
                var originalBounds = item.geometricBounds;
                var moveAmount = currentHeight - newHeight;

                // Adjust the height by shrinking from the bottom
                item.geometricBounds = [
                    originalBounds[0],  // Top edge remains the same
                    originalBounds[1],  // Left edge remains the same
                    originalBounds[2] - moveAmount,  // Move the bottom edge up to achieve the new height
                    originalBounds[3]   // Right edge remains the same
                ];

                // If it's a text frame or a rectangle, move the contents down after resizing the frame
                if (item instanceof TextFrame || item instanceof Rectangle || item instanceof GraphicLine) {
                    item.move(undefined, [0, moveAmount]); // Move the contents down after shrinking from the bottom
                }
            }
        }
    }

    // Function to process each item or group recursively
    function processItem(item) {
        if (item instanceof GraphicLine || item instanceof TextFrame || item instanceof Rectangle || item instanceof Image) {
            adjustHeight(item);
        } else if (item instanceof Group) {
            for (var j = 0; j < item.allPageItems.length; j++) {
                processItem(item.allPageItems[j]);
            }
        } else {
            // Continue processing without alerting
            // Commented out: alert("Selection contains unsupported items. Please select only vertical rules, text frames, picture boxes, or groups.");
        }
    }

    // Check if there are any selected objects
    if (app.selection.length === 0) {
        alert("Please select one or more vertical rules, text frames, picture boxes, or groups.");
    } else {
        // Iterate over all selected objects
        for (var i = 0; i < app.selection.length; i++) {
            processItem(app.selection[i]);
        }
    }
}, ScriptLanguage.JAVASCRIPT, null, UndoModes.ENTIRE_SCRIPT, "Adjust Item Height Based on Baseline Increment");
