// this is a test file, not for use within indesign 


app.doScript(function() {
    var doc = app.activeDocument;

    var gridPreferences = doc.gridPreferences;
    var baselineIncrement = gridPreferences.baselineDivision;
    var precisionThreshold = 0.1;

    var unsupportedItemsFound = false; // Track if unsupported items are found

    function adjustHeight(item) {
        var currentHeight = item.geometricBounds[2] - item.geometricBounds[0];
        if (currentHeight > baselineIncrement + precisionThreshold) {
            if (item instanceof Image && item.parent instanceof Rectangle) {
                adjustHeight(item.parent);
            } else {
                item.geometricBounds = [
                    item.geometricBounds[0],
                    item.geometricBounds[1],
                    item.geometricBounds[2] - baselineIncrement,
                    item.geometricBounds[3]
                ];
            }
        } else {
            var moveAmount = Math.abs(baselineIncrement);
            item.geometricBounds = [
                item.geometricBounds[0] - moveAmount,
                item.geometricBounds[1],
                item.geometricBounds[2] - moveAmount,
                item.geometricBounds[3]
            ];
        }
    }

    function findSmallerGroupItem(group) {
        var groupItem1 = group.pageItems[0];
        var groupItem2 = group.pageItems[1];

        var heightItem1 = groupItem1.geometricBounds[2] - groupItem1.geometricBounds[0];
        var heightItem2 = groupItem2.geometricBounds[2] - groupItem2.geometricBounds[0];

        var smallGroupItem, largeGroupItem;

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

    function resizeAndCenterSmallerGroupItem(smallGroupItem, largeGroupItem) {
        var largeTop = largeGroupItem.geometricBounds[0];
        var largeBottom = largeGroupItem.geometricBounds[2];
        var largeCenter = (largeTop + largeBottom) / 2;
        var baselineIncrementOffset = baselineIncrement / 2;

        var smallWidth = smallGroupItem.geometricBounds[3] - smallGroupItem.geometricBounds[1];
        var smallNewTop = largeCenter - (baselineIncrement / 2);

        smallGroupItem.geometricBounds = [
            smallNewTop,
            smallGroupItem.geometricBounds[1],
            smallNewTop + baselineIncrement / 2,
            smallGroupItem.geometricBounds[3]
        ];

        smallGroupItem.geometricBounds = [
            smallGroupItem.geometricBounds[0] - baselineIncrementOffset,
            smallGroupItem.geometricBounds[1],
            smallGroupItem.geometricBounds[2] - baselineIncrementOffset,
            smallGroupItem.geometricBounds[3]
        ];
    }

    function processItem(item) {
        if (item instanceof Group && item.pageItems.length === 2) {
            var groupItems = findSmallerGroupItem(item);
            resizeAndCenterSmallerGroupItem(groupItems.smallGroupItem, groupItems.largeGroupItem);
            adjustHeight(groupItems.largeGroupItem);
        } else if (item instanceof GraphicLine || item instanceof TextFrame || item instanceof Rectangle) {
            adjustHeight(item);
        } else if (item instanceof Group) {
            for (var j = 0; j < item.pageItems.length; j++) {
                processItem(item.pageItems[j]);
            }
        } else if (item instanceof Image) {
            var parent = item.parent;
            if (parent instanceof Rectangle) {
                adjustHeight(parent);
            } else {
                adjustHeight(item);
            }
        } else {
            unsupportedItemsFound = true; // Mark unsupported item found
        }
    }

    if (app.selection.length === 0) {
        alert("Please select one or more vertical rules, text frames, picture boxes, or groups.");
    } else {
        for (var i = 0; i < app.selection.length; i++) {
            processItem(app.selection[i]);
        }

        // If any unsupported items were found, show alert once
        if (unsupportedItemsFound) {
            alert("Some unsupported items were found in the selection. Please select only vertical rules, text frames, picture boxes, or groups.");
        }
    }
}, ScriptLanguage.JAVASCRIPT, null, UndoModes.ENTIRE_SCRIPT, "Adjust Item Height Based on Baseline Increment");
