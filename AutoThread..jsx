// NAME:AutoThread
// Status: Working  - works but improvements can be made regarding alert messages etc 
// FUNCTION - user selects last threaded text box and all other non threaded text boxes  - the script will then thread them left to right 
// Threading boxes is tedious as InDesign requires users to navigate to bottom right hand corner for each box they must thread - this script does it in one go
// OUTLINE - 
// Prefered keyboard shortcut is F7



//AutoThread
var gScriptName = "Thread Text Frames (Left to Right)";

try {
    app.doScript(main, undefined, undefined, UndoModes.ENTIRE_SCRIPT, gScriptName);
    // Force a screen refresh to update the threading display (if necessary).
    app.menuActions.itemByName("$ID/Force Redraw").invoke();
} catch (e) {
    alert(e + ". An error has occurred, try again.", gScriptName);
}

function main() {
    var selectedFrames = app.selection;

    if (selectedFrames.length < 2) {
        alert("Select at least two text frames to thread.", gScriptName);
        return;
    }

    var allThreaded = true;
    var threadedCount = 0;
    var firstFrameThreaded = false;

    // Determine if all selected frames are threaded
    for (var i = 0; i < selectedFrames.length; i++) {
        if (selectedFrames[i] instanceof TextFrame) {
            // Check if the frame has a previousTextFrame
            if (selectedFrames[i].previousTextFrame != null) {
                threadedCount++;
            } else if (selectedFrames[i].nextTextFrame != null) {
                // If it has a nextTextFrame, it is the first frame in a threading chain
                firstFrameThreaded = true;
                threadedCount++;
            } else {
                allThreaded = false;
            }
        }
    }

    // If all frames are threaded (including the first in a chain), show the appropriate message
    if (allThreaded && (threadedCount === selectedFrames.length)) {
        alert("All selected frames are already threaded. Please select frames that are not yet threaded.", gScriptName);
        return;
    }

    // Show error if more than one previously threaded frame is selected
    if (threadedCount > 1) {
        alert("Please ensure no more than one previously threaded text frame is part of your selection, then try again.", gScriptName);
        return;
    }

    // Sort the selected frames by their position
    selectedFrames.sort(compareFramesByPosition);

    // Add a carriage return to the end of each frame's text (except the first one)
    for (var i = 1; i < selectedFrames.length; i++) {
        addCarriageReturn(selectedFrames[i]);
    }

    // Start threading from the leftmost frame
    var leftmostFrame = selectedFrames[0];

    threadFrames(selectedFrames, leftmostFrame);
}

function compareFramesByPosition(frameA, frameB) {
    // Sort by x-position (geometricBounds[1] is the x-coordinate)
    return frameA.geometricBounds[1] - frameB.geometricBounds[1];
}

function threadFrames(frames, startFrame) {
    var previousFrame = startFrame;

    for (var i = 1; i < frames.length; i++) {
        var currentFrame = frames[i];

        // Thread the current frame to the previous one
        currentFrame.previousTextFrame = previousFrame;

        previousFrame = currentFrame;
    }
}

function addCarriageReturn(frame) {
    var text = frame.parentStory;
    var lastInsertionPoint = text.insertionPoints[-1];
    lastInsertionPoint.contents = "\r";
}
