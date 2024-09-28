// NAME:AutoThread
// STATUS: Ready
// FUNCTION: to thread a selection of text from left to right no matter the order a users selections
// Author: George Hannaford george.hannaford@telegraph.co.uk
// Prefered keyboard shortcut is alt+F3
// Version 1.0.0



//AutoThread
var gScriptName = "AutoThread";

try {
    app.doScript(main, undefined, undefined, UndoModes.ENTIRE_SCRIPT, gScriptName);
    app.menuActions.itemByName("$ID/Force Redraw").invoke();
} catch (e) {
    alert(e + ". An error has occurred, try again.", gScriptName);
}

function main() {
    var selectedFrames = app.selection;

    // Check if user has selected text frames
    if (selectedFrames.length < 2) {
        alert("Select at least two text frames to thread.", gScriptName);
        return;
    }

    // Validate that all selected items are text frames
    var textFrames = [];
    for (var i = 0; i < selectedFrames.length; i++) {
        if (selectedFrames[i] instanceof TextFrame) {
            textFrames.push(selectedFrames[i]);
        }
    }

    if (textFrames.length < 2) {
        alert("Please select at least two valid text frames.", gScriptName);
        return;
    }

    // Sort frames by horizontal position
    textFrames.sort(compareFramesByPosition);

    // Determine the starting frame of the thread
    var startingFrame = findStartingFrame(textFrames);

    // Determine the threading direction
    var threadingDirection = determineDirection(textFrames);

    // Re-threading logic
    try {
        rethreadFrames(textFrames, startingFrame, threadingDirection);
    } catch (e) {
        alert("Error during rethreading: " + e.message, gScriptName);
    }
}

function compareFramesByPosition(frameA, frameB) {
    // Sort by x-position (geometricBounds[1] is the x-coordinate)
    return frameA.geometricBounds[1] - frameB.geometricBounds[1];
}

function determineDirection(frames) {
    // Compare positions of the first and last frame
    var firstX = frames[0].geometricBounds[1];
    var lastX = frames[frames.length - 1].geometricBounds[1];
    return firstX < lastX ? "leftToRight" : "rightToLeft";
}

function findStartingFrame(frames) {
    // Check if any frame is a part of a thread
    var threadedFrames = [];
    
    // Helper function to check if an array contains an element
    function arrayContains(array, element) {
        for (var i = 0; i < array.length; i++) {
            if (array[i] === element) {
                return true;
            }
        }
        return false;
    }

    for (var i = 0; i < frames.length; i++) {
        var frame = frames[i];

        if (frame.previousTextFrame && !arrayContains(frames, frame.previousTextFrame)) {
            threadedFrames.push(frame.previousTextFrame);
        }
        if (frame.nextTextFrame && !arrayContains(frames, frame.nextTextFrame)) {
            threadedFrames.push(frame.nextTextFrame);
        }
    }

    if (threadedFrames.length > 0) {
        // Determine the farthest left (or right) frame
        var sortedThreadedFrames = threadedFrames.concat(frames).sort(compareFramesByPosition);
        return sortedThreadedFrames[0];
    }

    // If no threaded frames are found, return the first frame in the selection
    return frames[0];
}

function rethreadFrames(frames, startingFrame, direction) {
    var currentFrame = startingFrame;

    // Thread the frames based on the direction
    for (var i = 0; i < frames.length; i++) {
        if (direction === "leftToRight") {
            if (currentFrame.nextTextFrame !== frames[i] && currentFrame != frames[i]) {
                currentFrame.nextTextFrame = frames[i];
            }
            currentFrame = frames[i];
        } else if (direction === "rightToLeft") {
            if (currentFrame.previousTextFrame !== frames[i] && currentFrame != frames[i]) {
                currentFrame.previousTextFrame = frames[i];
            }
            currentFrame = frames[i];
        }
    }

}
