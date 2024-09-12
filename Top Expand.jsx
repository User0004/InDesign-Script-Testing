// NAME: Top Expand
// STATUS: Ready
// FUNCTION: Increase height of selected items from top one baseline for each time the script is run
// Best to use on non group grouped items for Nav. If needed use white arrow to handle groups, but black arrow is find for text boxes 
// Prefered keyboard shortcut is Top Expand F9


//Top Expand
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
        
        // Case 1: Shrink the height if it's greater than the baseline increment + threshold
        if (currentHeight > baselineIncrement + precisionThreshold) {
            // Reduce the bottom edge to shrink the height
            item.geometricBounds = [
                item.geometricBounds[0] - baselineIncrement, 
                item.geometricBounds[1], 
                item.geometricBounds[2], 
                item.geometricBounds[3]
            ];
        } 
        // Case 2: Move the item upwards if its height is equal to or less than one baseline (within threshold)
        else {
            var moveAmount = Math.abs(baselineIncrement); // Ensure positive value

            // Move the object upwards by one baseline increment without changing its height
            item.geometricBounds = [
                item.geometricBounds[0] - moveAmount,
                item.geometricBounds[1],
                item.geometricBounds[2] - moveAmount,
                item.geometricBounds[3]
            ];
        }
    }

    // Function to process each item or group recursively
    function processItem(item) {
        if (item instanceof GraphicLine || item instanceof TextFrame || item instanceof Rectangle) {
            adjustHeight(item);
        } else if (item instanceof Group) {
            for (var j = 0; j < item.allPageItems.length; j++) {
                processItem(item.allPageItems[j]);
            }
        } else if (item instanceof Image) {
            // Check if the image is inside a rectangle (frame)
            var parent = item.parent;
            if (parent instanceof Rectangle) {
                adjustHeight(parent); // Adjust the frame's height, not the image itself
            } else {
                // For standalone images not inside a rectangle, adjust the image's bounds directly
                adjustHeight(item);
            }
        } else {
            alert("Selection contains unsupported items. Please select only vertical rules, text frames, picture boxes, or groups.");
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
