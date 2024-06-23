var myFolder = Folder.selectDialog("Select the folder with images");
if (myFolder != null) {
    var myFiles = myFolder.getFiles("*.*");
    var myDocument = app.activeDocument;
    
    for (var i = 0; i < myFiles.length; i++) {
        var myPage = myDocument.pages.item(i);
        var myRectangle = myPage.rectangles.add();
        myRectangle.geometricBounds = ["12.7mm", "12.7mm", "150mm", "200mm"];
        myRectangle.place(myFiles[i]);
        myRectangle.fit(FitOptions.FILL_PROPORTIONALLY);
    }
}
