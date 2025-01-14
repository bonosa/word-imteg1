async function setBackgroundImage(imageUrl) {
  try {
    const blob = await fetchImageBlob(imageUrl);
    const base64Image = await blobToBase64(blob);

    await Word.run(async (context) => {
      const body = context.document.body;

      // Insert the image inline
      const inlinePicture = body.insertInlinePictureFromBase64(
        base64Image,
        Word.InsertLocation.start
      );

      // Convert the inline image to a floating image
      inlinePicture.convertToFloatingImage();

      // Set the floating image properties
      inlinePicture.floatingFormat.wrapText = "behindText"; // Place behind text
      inlinePicture.floatingFormat.horizontalPositionAlignment = Word.Alignment.center; // Center horizontally
      inlinePicture.floatingFormat.verticalPositionAlignment = Word.VerticalAlignment.center; // Center vertically

      // Sync changes
      await context.sync();
      console.log("Background image set successfully.");
    });
  } catch (error) {
    console.error("Error setting background image:", error);
  }
}