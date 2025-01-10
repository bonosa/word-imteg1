/* taskpane.js */

// Replace with your actual Unsplash Access Key
const UNSPLASH_ACCESS_KEY = "YOUR_UNSPLASH_ACCESS_KEY";

Office.onReady(() => {
  // Once Office.js is ready, attach event listeners
  document.getElementById("searchButton").addEventListener("click", onSearch);
});

/**
 * Handler for the search button
 */
async function onSearch() {
  const query = document.getElementById("searchQuery").value.trim();
  if (!query) return;

  const results = await fetchImagesFromUnsplash(query);
  displayImages(results);
}

/**
 * Call Unsplash API to get images for the given query
 */
async function fetchImagesFromUnsplash(query) {
  const url = `https://api.unsplash.com/search/photos?client_id=${UNSPLASH_ACCESS_KEY}&query=${encodeURIComponent(query)}`;
  try {
    const response = await fetch(url);
    const data = await response.json();
    return data.results || [];
  } catch (err) {
    console.error("Error fetching from Unsplash:", err);
    return [];
  }
}

/**
 * Display the images in our gallery. Each image is clickable.
 */
function displayImages(images) {
  const gallery = document.getElementById("gallery");
  gallery.innerHTML = "";

  images.forEach((img) => {
    const container = document.createElement("div");

    // Create an <img> for the thumbnail
    const thumbnail = document.createElement("img");
    thumbnail.src = img.urls.thumb;
    thumbnail.title = `Photo by ${img.user.name} on Unsplash`;
    thumbnail.addEventListener("click", () => insertImageIntoDoc(img.urls.full));

    // Create a <p> for attribution
    const caption = document.createElement("p");
    caption.innerHTML = `Photo by ${img.user.name} on <a href="${img.links.html}" target="_blank">Unsplash</a>`;

    container.appendChild(thumbnail);
    container.appendChild(caption);
    gallery.appendChild(container);
  });
}

/**
 * Insert the selected image into the Word document
 */
async function insertImageIntoDoc(imageUrl) {
  try {
    const blob = await fetchImageBlob(imageUrl);
    const base64Image = await blobToBase64(blob);

    await Word.run(async (context) => {
      const docBody = context.document.body;
      docBody.insertInlinePictureFromBase64(base64Image, Word.InsertLocation.end);
      await context.sync();
    });
  } catch (error) {
    console.error("Error inserting image:", error);
  }
}

/**
 * Fetch the image as a Blob
 */
async function fetchImageBlob(url) {
  const response = await fetch(url);
  return response.blob();
}

/**
 * Convert a Blob to a base64 string
 */
function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => {
      // DataURL -> we split after comma to get the base64 part
      const base64String = reader.result.split(",")[1];
      resolve(base64String);
    };
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}
