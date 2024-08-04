// Get the active document
var doc = app.activeDocument;

// Get the active page
var activePage = doc.layoutWindows[0].activePage;

// Get the bounds of the active page
var pageBounds = activePage.bounds; // [y1, x1, y2, x2]

// Calculate the center of the page
var pageCenterX = (pageBounds[1] + pageBounds[3]) / 2;
var pageCenterY = (pageBounds[0] + pageBounds[2]) / 2;

// Get all the frames on the active page
var frames = activePage.allPageItems;

// Loop through each frame and center it
for (var i = 0; i < frames.length; i++) {
    var frame = frames[i];

    // Get the bounds of the frame
    var frameBounds = frame.geometricBounds; // [y1, x1, y2, x2]

    // Calculate the center of the frame
    var frameWidth = frameBounds[3] - frameBounds[1];
    var frameHeight = frameBounds[2] - frameBounds[0];

    // Calculate the new bounds for the frame
    var newBounds = [
        pageCenterY - frameHeight / 2, // y1
        pageCenterX - frameWidth / 2,  // x1
        pageCenterY + frameHeight / 2, // y2
        pageCenterX + frameWidth / 2   // x2
    ];

    // Set the new bounds for the frame
    frame.geometricBounds = newBounds;
}

alert("Frames centered on the page!");
