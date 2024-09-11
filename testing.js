// work in progrees improvement to shrink scritps 

//shrink from bottom
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

        if (currentHeight > baselineIncrement + precisionThreshold) {
            // Reduce the bottom edge to shrink the height
            if (item instanceof Image && item.parent instanceof Rectangle) {
                adjustHeight(item.parent); // Adjust the height of the parent frame, not the image itself
            } else {
                item.geometricBounds = [
                    item.geometricBounds[0],
                    item.geometricBounds[1], 
                    item.geometricBounds[2] - baselineIncrement,
                    item.geometricBounds[3]
                ];
            }
        } else {
            // Move the item upwards if its height is equal to or less than one baseline (within threshold)
            var moveAmount = Math.abs(baselineIncrement); // Ensure positive value

            item.geometricBounds = [
                item.geometricBounds[0] - moveAmount,
                item.geometricBounds[1],
                item.geometricBounds[2] - moveAmount,
                item.geometricBounds[3]
            ];
        }
    }

    // Function to find and assign smaller and larger items in a group of two
    function findSmallerGroupItem(group) {
        var groupItem1 = group.allPageItems[0];
        var groupItem2 = group.allPageItems[1];

        var heightItem1 = groupItem1.geometricBounds[2] - groupItem1.geometricBounds[0];
        var heightItem2 = groupItem2.geometricBounds[2] - groupItem2.geometricBounds[0];

        var smallGroupItem, largeGroupItem;

        // Compare heights and assign smaller and larger
        if (heightItem1 < heightItem2) {
            smallGroupItem = groupItem1;
            largeGroupItem = groupItem2;
        } else {
            smallGroupItem = groupItem2;
            largeGroupItem = groupItem1;
        }

        return {
            smallGroupItem: smallGroupItem,
            largeGroupItem: largeGroupItem
        };
    }

    // Function to ensure smaller item is centered within the larger item
    function ensureSmallerGroupItemIsCentre(smallGroupItem, largeGroupItem) {
        var largeTop = largeGroupItem.geometricBounds[0];
        var largeBottom = largeGroupItem.geometricBounds[2];
        var largeCenter = (largeTop + largeBottom) / 2;

        var smallHeight = smallGroupItem.geometricBounds[2] - smallGroupItem.geometricBounds[0];
        var smallTop = largeCenter - (smallHeight / 2);
       
        // Adjust the smaller item’s vertical position to be centered within the larger one
        smallGroupItem.geometricBounds = [
            smallTop,
            smallGroupItem.geometricBounds[1],
            smallTop + smallHeight,
            smallGroupItem.geometricBounds[3]
        ];
    }

    // Function to resize the smaller group item to match baseline increment
    function smallerGroupItemSize(smallGroupItem) {
        smallGroupItem.geometricBounds = [
            smallGroupItem.geometricBounds[0] + baselineIncrement ,
            smallGroupItem.geometricBounds[1],
            smallGroupItem.geometricBounds[0],
            smallGroupItem.geometricBounds[3]
        ];
    }

    // Function to offset the position of the smaller group item
    function offsetSmallGroupItem(smallGroupItem) {
        var offsetAmount = baselineIncrement / 2;
        var currentBounds = smallGroupItem.geometricBounds;

        smallGroupItem.geometricBounds = [
            currentBounds[0] + offsetAmount,
            currentBounds[1],
            currentBounds[2] + offsetAmount,
            currentBounds[3]
        ];
    }

    // Function to process each item or group recursively
    function processItem(item) {
        if (item instanceof Group && item.allPageItems.length === 2) {
            // If the item is a group with exactly two items, find smaller and larger items
            var groupItems = findSmallerGroupItem(item);

            // Ensure the smaller item is centered within the larger item
            ensureSmallerGroupItemIsCentre(groupItems.smallGroupItem, groupItems.largeGroupItem);

            // Resize the smaller group item
            smallerGroupItemSize(groupItems.smallGroupItem);

            // Offset the smaller group item
            offsetSmallGroupItem(groupItems.smallGroupItem);

            // Now adjust the height of the larger item as well
            adjustHeight(groupItems.largeGroupItem);

        } else if (item instanceof GraphicLine || item instanceof TextFrame || item instanceof Rectangle) {
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
