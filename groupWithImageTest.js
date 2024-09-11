// centre test code here 




app.doScript(function() {
    // Function to handle centering of two items (whether grouped or not)
    function centerTwoItems(item1, item2) {
        var bounds1 = item1.geometricBounds;
        var bounds2 = item2.geometricBounds;

        var height1 = bounds1[2] - bounds1[0]; // Bottom - Top = Height
        var height2 = bounds2[2] - bounds2[0]; // Bottom - Top = Height

        var smallerItem, largerItem;

        // Determine which item is smaller
        if (height1 < height2) {
            smallerItem = item1;
            largerItem = item2;
        } else {
            smallerItem = item2;
            largerItem = item1;
        }

        // Get the center of the larger item
        var largeTop = largerItem.geometricBounds[0];
        var largeBottom = largerItem.geometricBounds[2];
        var largeCenter = (largeTop + largeBottom) / 2;

        // Get the height of the smaller item
        var smallHeight = smallerItem.geometricBounds[2] - smallerItem.geometricBounds[0];

        // Calculate the new top position for the smaller item so it's centered
        var newSmallTop = largeCenter - (smallHeight / 2);
        var newSmallBottom = newSmallTop + smallHeight;

        // Set the new bounds for the smaller item (vertical centering only)
        smallerItem.geometricBounds = [
            newSmallTop,                  // New top
            smallerItem.geometricBounds[1], // Keep left unchanged
            newSmallBottom,               // New bottom
            smallerItem.geometricBounds[3]  // Keep right unchanged
        ];
    }

    // Main function to process the selection
    function processSelection(selection) {
        if (selection.length === 1 && selection[0] instanceof Group) {
            // If the selection is a group, check if it has exactly two items
            var groupItems = selection[0].allPageItems;
            if (groupItems.length === 2) {
                centerTwoItems(groupItems[0], groupItems[1]); // Center the two items in the group
            } else {
                alert("Group must contain exactly two items.");
            }
        } else if (selection.length === 2) {
            // If exactly two items are selected (not a group), center them
            centerTwoItems(selection[0], selection[1]);
        } else {
            alert("Please select exactly two objects or a group containing two objects.");
        }
    }

    // Check the user's selection
    var selection = app.selection;

    // If there is no selection or more than two selected items, show an alert
    if (selection.length === 0) {
        alert("Please select exactly two objects or a group containing two objects.");
    } else {
        // Process the user's selection
        processSelection(selection);
    }

}, ScriptLanguage.JAVASCRIPT, null, UndoModes.ENTIRE_SCRIPT, "Center Smaller Item Vertically (Group Support)");
