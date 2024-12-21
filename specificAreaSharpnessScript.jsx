#target photoshop

// Function to unlock the background layer (if needed)
function unlockLayer(layer) {
    if (layer.isBackgroundLayer) {
        layer.isBackgroundLayer = false; // Convert background to an editable layer
    }
}

// Function to calculate sharpness of a section
function calculateSharpnessForRegion(imageDoc, x, y, width, height) {
    // Ensure dimensions are integers
    x = Math.round(x);
    y = Math.round(y);
    width = Math.round(width);
    height = Math.round(height);

    // Skip regions with invalid dimensions
    if (width <= 0 || height <= 0 || x + width > imageDoc.width || y + height > imageDoc.height) {
        return 0; // Return 0 sharpness for invalid regions
    }

    // Select the specific region
    try {
        imageDoc.selection.select([
            [x, y],
            [x + width, y],
            [x + width, y + height],
            [x, y + height]
        ]);
        imageDoc.selection.copy();

        // Create a new temporary document for the region
        var tempDoc = app.documents.add(width, height, imageDoc.resolution, "TempDoc", NewDocumentMode.GRAYSCALE);
        tempDoc.paste();

        // Apply High Pass filter
        tempDoc.activeLayer.applyHighPass(10);

        // Calculate sharpness based on pixel histogram
        var sharpnessScore = 0;
        var pixels = tempDoc.channels[0].histogram;
        for (var i = 0; i < pixels.length; i++) {
            sharpnessScore += pixels[i] * i; // Weighted sum
        }

        // Close the temporary document
        tempDoc.close(SaveOptions.DONOTSAVECHANGES);

        return sharpnessScore;

    } catch (err) {
        // If selection or copy fails, return 0 sharpness
        return 0;
    }
}

// Function to find the sharpest section in the image
function findSharpestSection(imageDoc, gridSize) {
    var width = imageDoc.width.as("px");
    var height = imageDoc.height.as("px");

    var regionWidth = Math.floor(width / gridSize);
    var regionHeight = Math.floor(height / gridSize);

    var maxSharpness = 0;

    // Loop through the grid
    for (var i = 0; i < gridSize; i++) {
        for (var j = 0; j < gridSize; j++) {
            var x = i * regionWidth;
            var y = j * regionHeight;

            // Calculate sharpness for this section
            var sharpness = calculateSharpnessForRegion(imageDoc, x, y, regionWidth, regionHeight);

            // Track the maximum sharpness
            if (sharpness > maxSharpness) {
                maxSharpness = sharpness;
            }
        }
    }

    return maxSharpness; // Return the highest sharpness value from all sections
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

    var gridSize = 8; // Divide image into a 4x4 grid

    // Process each file
    for (var i = 0; i < files.length; i++) {
        var file = files[i];
        var imageDoc = null;

        try {
            imageDoc = open(file); // Open the image file

            // Unlock the background layer to ensure it's editable
            unlockLayer(imageDoc.activeLayer);

            // Find the sharpest section in the image
            var sharpness = findSharpestSection(imageDoc, gridSize);
            sharpnessScores.push({ file: file, score: sharpness });

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

    // Assign ranks: sharpest is 1, least sharp is total number of files
    var totalFiles = sharpnessScores.length;
    for (var j = 0; j < sharpnessScores.length; j++) {
        sharpnessScores[j].ranking = j + 1; // Rank in ascending order (1 to totalFiles)
    }

    // Display results
    var result = "Ranking: Smaller numbers represent sharper images (based on sharpest section).\n\nSharpness Rankings (1 to " + totalFiles + "):\n";
    for (var k = 0; k < sharpnessScores.length; k++) {
        result += sharpnessScores[k].file.name + ": " + sharpnessScores[k].ranking + "\n";
    }

    // Save the top 5 sharpest images
    var topCount = Math.min(5, sharpnessScores.length);
    for (var l = 0; l < topCount; l++) {
        var topImage = open(sharpnessScores[l].file);
        saveCopy(topImage, outputFolder, "Top_" + (l + 1) + "_" + sharpnessScores[l].file.name);
        topImage.close(SaveOptions.DONOTSAVECHANGES);
    }

    result += "\nTop 5 sharpest images saved in the OUTPUT folder.";

    alert(result); // Display the result
}

// Execute the main function
findSharpestImages();

