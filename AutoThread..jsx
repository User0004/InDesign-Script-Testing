// NAME:AutoThread
// Status: Working  - works but improvements can be made regarding alert messages etc 
// FUNCTION - user selects last threaded text box and all other non threaded text boxes  - the script will then thread them left to right 
// Threading boxes is tedious as InDesign requires users to navigate to bottom right hand corner for each box they must thread - this script does it in one go
// OUTLINE - 
// Prefered keyboard shortcut is F7

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

    // Check for threaded frames
    var threadedCount = 0;
    for (var i = 0; i < selectedFrames.length; i++) {
        if (selectedFrames[i] instanceof TextFrame && selectedFrames[i].previousTextFrame != null) {
            threadedCount++;
        }
    }

    // Show error if more than one previously threaded frame is selected
    if (threadedCount > 1) {
        alert("Please ensure no more than one previously threaded text frame is selected.", gScriptName);
        return;
    }

    // Sort the selected frames by their position
    selectedFrames.sort(compareFramesByPosition);

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
