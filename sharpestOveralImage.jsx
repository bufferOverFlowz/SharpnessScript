#target photoshop

// Function to calculate sharpness score
function calculateSharpness(imageDoc) {
    var tempDoc = imageDoc.duplicate("TempDoc", true); // Duplicate document
    tempDoc.changeMode(ChangeMode.GRAYSCALE);         // Convert to grayscale
    tempDoc.activeLayer.applyHighPass(10);           // Apply High Pass filter
    tempDoc.flatten();                               // Flatten layers

    // Calculate sharpness based on pixel histogram
    var sharpnessScore = 0;
    var pixels = tempDoc.channels[0].histogram;
    for (var i = 0; i < pixels.length; i++) {
        sharpnessScore += pixels[i] * i; // Weighted sum
    }

    tempDoc.close(SaveOptions.DONOTSAVECHANGES);    // Close temporary document
    return sharpnessScore;
}

// Function to save a copy of the image
function saveCopy(imageDoc, outputFolder, fileName) {
    var saveOptions = new JPEGSaveOptions();
    saveOptions.quality = 12; // Maximum quality
    var savePath = new File(outputFolder + "/" + fileName);
    imageDoc.saveAs(savePath, saveOptions, true, Extension.LOWERCASE);
}

// Main function to find sharpest images
function findSharpestImages() {
    // Prompt the user to select a folder
    var folder = Folder.selectDialog("Select a folder containing images");
    if (!folder) {
        alert("No folder selected. Please try again.");
        return;
    }

    // Get all valid image files in the folder
    var files = folder.getFiles(function(file) {
        return file instanceof File && /\.(jpg|jpeg|png|tif|tiff|bmp|gif|psd)$/i.test(file.name);
    });

    if (files.length === 0) {
        alert("No valid image files found in the selected folder.\nSupported formats: JPG, JPEG, PNG, TIF, TIFF, BMP, GIF, PSD.");
        return;
    }

    var sharpnessScores = [];
    var outputFolder = new Folder(folder + "/OUTPUT");

    // Create OUTPUT folder if it doesn't exist
    if (!outputFolder.exists) {
        outputFolder.create();
    }

    // Process each file
    for (var i = 0; i < files.length; i++) {
        var file = files[i];
        var imageDoc = null;

        try {
            imageDoc = open(file); // Open the image file

            // Calculate sharpness score
            var score = Math.round(calculateSharpness(imageDoc)); // Round for readability
            sharpnessScores.push({ file: file, doc: imageDoc, score: score });

        } catch (err) {
            alert("Error processing file: " + file.name + "\nError: " + err.message);
        } finally {
            if (imageDoc) {
                imageDoc.close(SaveOptions.DONOTSAVECHANGES); // Ensure file is closed
            }
        }
    }

    // Sort results by sharpness score (descending order)
    sharpnessScores.sort(function(a, b) {
        return b.score - a.score;
    });

    // Display results
    var result = "Sharpness Ratings:\n";
    for (var j = 0; j < sharpnessScores.length; j++) {
        result += sharpnessScores[j].file.name + ": " + sharpnessScores[j].score + "\n";
    }

    // Save the top 5 sharpest images
    var topCount = Math.min(5, sharpnessScores.length);
    for (var k = 0; k < topCount; k++) {
        var topImage = open(sharpnessScores[k].file);
        saveCopy(topImage, outputFolder, "Top_" + (k + 1) + "_" + sharpnessScores[k].file.name);
        topImage.close(SaveOptions.DONOTSAVECHANGES);
    }

    result += "\nTop 5 sharpest images saved in the OUTPUT folder.";

    alert(result); // Display the result
}

// Execute the main function
findSharpestImages();

