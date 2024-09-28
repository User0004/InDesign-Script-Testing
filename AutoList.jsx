// NAME:AutoList
// STATUS: Ready 
// FUNCTION: To search for an array of words which is under 10 in length not contaning a period (with exceptions) and apply the first instance    subheader throughout following text
// The script allows the designer to have many stacked subheaders of differnt styles. However the chain of styles is broken upon a hard enter
// This is a sister script to AutoSubheader ---- this script works where there is bodycopy paragraphs and many  
// Updates: There are hard coded styles, would be good to improve it by making a temp parageaph style 
// Author: George Hannaford george.hannaford@telegraph.co.uk
// Prefered keyboard shortcut is alt+F6
// Version 1.0.0

//AutoList
function findAndApplyTextSettings() {
    // Ensure that there's an active document open
    if (app.documents.length > 0) {
        var doc = app.activeDocument;
        var selection = app.selection;

        // Ensure there's a valid selection
        if (selection.length > 0 && (selection[0] instanceof TextFrame || selection[0].hasOwnProperty('insertionPoints'))) {
            var selectedTextFrames = [];
            if (selection[0] instanceof TextFrame) {
                selectedTextFrames.push(selection[0]);
            } else if (selection[0].hasOwnProperty('insertionPoints')) {
                selectedTextFrames.push(selection[0].parentTextFrames[0]);
            }

            var foundTargetParagraph = false;
            var targetTextSettingsList = [];

            // Iterate over each selected text frame
            for (var j = 0; j < selectedTextFrames.length; j++) {
                var paragraphs = selectedTextFrames[j].parentStory.paragraphs.everyItem().getElements();

                // Find the first suitable paragraph and all consecutive ones
                for (var i = 0; i < paragraphs.length; i++) {
                    var consecutiveParagraphs = [];
                    var allConsecutiveSettings = [];

                    for (var k = i; k < paragraphs.length; k++) {
                        var para = paragraphs[k];
                        var words = para.words.length;

                        // Check updated criteria for paragraphs
                        if (words > 0 && words <= 10) {
                            var lastWord = para.words.lastItem();
                            var lastWordContents = lastWord.contents;

                            // Check if the last word ends with a period
                            if (lastWordContents.charAt(lastWordContents.length - 1) === ".") {
                                // Check if the period is part of a number or abbreviation
                                var periodPattern = /\d+\.\d+|[a-zA-Z]\.\s*$/;
                                if (!periodPattern.test(lastWordContents)) {
                                    consecutiveParagraphs.push(para);
                                    allConsecutiveSettings.push(getTextSettings(para));
                                }
                            } else {
                                consecutiveParagraphs.push(para);
                                allConsecutiveSettings.push(getTextSettings(para));
                            }
                        } else {
                            break; // Stop checking if the paragraph doesn't meet criteria
                        }
                    }

                    if (consecutiveParagraphs.length > 0) {
                        targetTextSettingsList = allConsecutiveSettings;
                        foundTargetParagraph = true;
                        break; // Exit loop once found
                    }
                }

                if (foundTargetParagraph) {
                    // Apply text settings to similar paragraphs within an undoable action
                    app.doScript(function() {
                        var currentConsecutiveIndex = 0;
                        for (var i = 0; i < paragraphs.length; i++) {
                            var para = paragraphs[i];
                            var words = para.words.length;

                            // Apply text settings if it's a similar paragraph
                            if (words > 0 && words <= 10) {
                                var lastWord = para.words.lastItem();
                                var lastWordContents = lastWord.contents;

                                if (lastWordContents.charAt(lastWordContents.length - 1) === ".") {
                                    var periodPattern = /\d+\.\d+|[a-zA-Z]\.\s*$/;
                                    if (!periodPattern.test(lastWordContents)) {
                                        if (currentConsecutiveIndex < targetTextSettingsList.length) {
                                            applyTextSettings(para, targetTextSettingsList[currentConsecutiveIndex]);
                                            currentConsecutiveIndex++;
                                        }
                                    }
                                } else {
                                    if (currentConsecutiveIndex < targetTextSettingsList.length) {
                                        applyTextSettings(para, targetTextSettingsList[currentConsecutiveIndex]);
                                        currentConsecutiveIndex++;
                                    }
                                }
                            } else {
                                currentConsecutiveIndex = 0; // Reset if paragraph does not meet criteria
                            }
                        }
                    }, ScriptLanguage.JAVASCRIPT, [], UndoModes.ENTIRE_SCRIPT, "Apply Text Settings");
                } else {
                    alert("No suitable paragraphs found in the selected text frame.");
                }
            }
        } else {
            alert("Please select a text frame or place the cursor inside a text frame.");
        }
    } else {
        alert("No active document found.");
    }
}

// Function to get text settings from a paragraph
function getTextSettings(paragraph) {
    var settings = {};

    var paragraphStyle = paragraph.appliedParagraphStyle;
    if (paragraphStyle.isValid && isStyleUnmodified(paragraph, paragraphStyle)) {
        settings.paragraphStyle = paragraphStyle;
    } else {
        // Fallback: Gather all the settings manually
        settings.appliedFont = paragraph.appliedFont;
        settings.pointSize = paragraph.pointSize;
        settings.fillColor = paragraph.fillColor;
        settings.fillTint = paragraph.fillTint;
        settings.justification = paragraph.justification;
        settings.leading = paragraph.leading;
        settings.firstLineIndent = paragraph.firstLineIndent;
        settings.baselineShift = paragraph.baselineShift;
        settings.tracking = paragraph.tracking;
        settings.capitalization = getCapitalization(paragraph.capitalization);
        settings.alignToBaseline = paragraph.alignToBaseline;
        settings.ruleAbove = paragraph.ruleAbove === true;
        settings.ruleBelow = paragraph.ruleBelow === true;

        // Rule Above settings
        settings.ruleAboveColor = paragraph.ruleAboveColor;
        settings.ruleAboveGapColor = paragraph.ruleAboveGapColor;
        settings.ruleAboveGapOverprint = paragraph.ruleAboveGapOverprint === true;
        settings.ruleAboveGapTint = paragraph.ruleAboveGapTint;
        settings.ruleAboveLeftIndent = paragraph.ruleAboveLeftIndent;
        settings.ruleAboveLineWeight = paragraph.ruleAboveLineWeight;
        settings.ruleAboveOffset = paragraph.ruleAboveOffset;
        settings.ruleAboveOverprint = paragraph.ruleAboveOverprint === true;
        settings.ruleAboveRightIndent = paragraph.ruleAboveRightIndent;
        settings.ruleAboveTint = paragraph.ruleAboveTint;
        settings.ruleAboveType = paragraph.ruleAboveType;
        settings.ruleAboveWidth = paragraph.ruleAboveWidth;

        // Rule Below settings
        settings.ruleBelowColor = paragraph.ruleBelowColor;
        settings.ruleBelowGapColor = paragraph.ruleBelowGapColor;
        settings.ruleBelowGapOverprint = paragraph.ruleBelowGapOverprint === true;
        settings.ruleBelowGapTint = paragraph.ruleBelowGapTint;
        settings.ruleBelowLeftIndent = paragraph.ruleBelowLeftIndent;
        settings.ruleBelowLineWeight = paragraph.ruleBelowLineWeight;
        settings.ruleBelowOffset = paragraph.ruleBelowOffset;
        settings.ruleBelowOverprint = paragraph.ruleBelowOverprint === true;
        settings.ruleBelowRightIndent = paragraph.ruleBelowRightIndent;
        settings.ruleBelowTint = paragraph.ruleBelowTint;
        settings.ruleBelowType = paragraph.ruleBelowType;
        settings.ruleBelowWidth = paragraph.ruleBelowWidth;

        // Additional paragraph settings
        settings.spaceBefore = paragraph.spaceBefore;
        settings.spaceAfter = paragraph.spaceAfter;
        settings.hyphenation = paragraph.hyphenation;
    }

    return settings;
}

// Function to apply text settings to a paragraph
function applyTextSettings(paragraph, settings) {
    if (settings.paragraphStyle) {
        // Apply paragraph style if it exists and is unmodified
        paragraph.applyParagraphStyle(settings.paragraphStyle, true);
    } else {
        // Apply individual settings if paragraph style is not available
        paragraph.appliedFont = settings.appliedFont;
        paragraph.pointSize = settings.pointSize;
        paragraph.fillColor = settings.fillColor;
        paragraph.fillTint = settings.fillTint;
        paragraph.leading = settings.leading;
        paragraph.firstLineIndent = settings.firstLineIndent;
        paragraph.justification = settings.justification;
        paragraph.baselineShift = settings.baselineShift;
        paragraph.tracking = settings.tracking;
        paragraph.capitalization = settings.capitalization;
        paragraph.alignToBaseline = settings.alignToBaseline;
        paragraph.ruleAbove = settings.ruleAbove;
        paragraph.ruleBelow = settings.ruleBelow;

        // Apply rule above settings
        paragraph.ruleAboveColor = settings.ruleAboveColor;
        paragraph.ruleAboveGapColor = settings.ruleAboveGapColor;
        paragraph.ruleAboveGapOverprint = settings.ruleAboveGapOverprint;
        paragraph.ruleAboveGapTint = settings.ruleAboveGapTint;
        paragraph.ruleAboveLeftIndent = settings.ruleAboveLeftIndent;
        paragraph.ruleAboveLineWeight = settings.ruleAboveLineWeight;
        paragraph.ruleAboveOffset = settings.ruleAboveOffset;
        paragraph.ruleAboveOverprint = settings.ruleAboveOverprint;
        paragraph.ruleAboveRightIndent = settings.ruleAboveRightIndent;
        paragraph.ruleAboveTint = settings.ruleAboveTint;
        paragraph.ruleAboveType = settings.ruleAboveType;
        paragraph.ruleAboveWidth = settings.ruleAboveWidth;

        // Apply rule below settings
        paragraph.ruleBelowColor = settings.ruleBelowColor;
        paragraph.ruleBelowGapColor = settings.ruleBelowGapColor;
        paragraph.ruleBelowGapOverprint = settings.ruleBelowGapOverprint;
        paragraph.ruleBelowGapTint = settings.ruleBelowGapTint;
        paragraph.ruleBelowLeftIndent = settings.ruleBelowLeftIndent;
        paragraph.ruleBelowLineWeight = settings.ruleBelowLineWeight;
        paragraph.ruleBelowOffset = settings.ruleBelowOffset;
        paragraph.ruleBelowOverprint = settings.ruleBelowOverprint;
        paragraph.ruleBelowRightIndent = settings.ruleBelowRightIndent;
        paragraph.ruleBelowTint = settings.ruleBelowTint;
        paragraph.ruleBelowType = settings.ruleBelowType;
        paragraph.ruleBelowWidth = settings.ruleBelowWidth;

        // Additional paragraph settings
        paragraph.spaceBefore = settings.spaceBefore;
        paragraph.spaceAfter = settings.spaceAfter;
        paragraph.hyphenation = settings.hyphenation;

        // Apply word-level settings to each word in the paragraph
        var paraWords = paragraph.words.everyItem().getElements();
        for (var k = 0; k < paraWords.length; k++) {
            paraWords[k].appliedFont = settings.appliedFont;
            paraWords[k].pointSize = settings.pointSize;
            paraWords[k].fillColor = settings.fillColor;
            paraWords[k].fillTint = settings.fillTint;
            paraWords[k].leading = settings.leading;
            paraWords[k].baselineShift = settings.baselineShift;
            paraWords[k].tracking = settings.tracking;
            paraWords[k].capitalization = settings.capitalization;
            paraWords[k].alignToBaseline = settings.alignToBaseline;
        }
    }
}

// Function to check if the style is unmodified (compare paragraph properties to the style)
function isStyleUnmodified(paragraph, paragraphStyle) {
    return paragraph.appliedFont === paragraphStyle.appliedFont &&
           paragraph.pointSize === paragraphStyle.pointSize &&
           paragraph.fillColor === paragraphStyle.fillColor &&
           paragraph.leading === paragraphStyle.leading &&
           paragraph.firstLineIndent === paragraphStyle.firstLineIndent;
}

// Helper function to map capitalization constants to InDesign values
function getCapitalization(constantValue) {
    switch (constantValue) {
        case 1634493296:
            return Capitalization.ALL_CAPS;
        case 1664250723:
            return Capitalization.CAP_TO_SMALL_CAP;
        case 1852797549:
            return Capitalization.NORMAL;
        case 1936548720:
            return Capitalization.SMALL_CAPS;
        default:
            return Capitalization.NORMAL; // Default to normal capitalization
    }
}

// Run the function
findAndApplyTextSettings();
