/* taskpane.js */
const UNSPLASH_ACCESS_KEY = "n3iOXASl_P5newMXNlMH8ny_ZbLXAdnHPuMum5oq6ls";

let insertedImagesMap = {};

Office.onReady(() => {
  const searchButton = document.getElementById("searchButton");
  searchButton.addEventListener("click", onSearch);
});

/**
 * When user clicks "Search", fetch 10 new images and replace the old gallery
 */
async function onSearch() {
  const query = document.getElementById("searchQuery").value.trim();
  if (!query) return;

  document.getElementById("gallery").innerHTML = "";

  const results = await fetchImagesFromUnsplash(query, 10); // Updated to fetch 10 images
  displayImages(results);
}

/**
 * Fetch images from Unsplash
 */
async function fetchImagesFromUnsplash(query, count) {
  const url = `https://api.unsplash.com/search/photos?client_id=${UNSPLASH_ACCESS_KEY}&query=${encodeURIComponent(query)}&per_page=${count}`;
  try {
    const response = await fetch(url);
    const data = await response.json();
    return data.results || [];
  } catch (err) {
    console.error("Error fetching Unsplash:", err);
    return [];
  }
}

/**
 * Display the fetched images in the gallery
 */
function displayImages(images) {
  const gallery = document.getElementById("gallery");
  images.forEach((img) => {
    const container = document.createElement("div");

    const thumbnail = document.createElement("img");
    thumbnail.src = img.urls.thumb;
    thumbnail.title = `Photo by ${img.user.name} on Unsplash`;
    thumbnail.addEventListener("click", () => onImageClick(img.urls.full));

    const caption = document.createElement("p");
    caption.innerHTML = `Photo by ${img.user.name} on <a href="${img.links.html}" target="_blank">Unsplash</a>`;

    container.appendChild(thumbnail);
    container.appendChild(caption);
    gallery.appendChild(container);
  });
}

/**
 * Set the clicked image as the background of the current page
 */
async function onImageClick(fullUrl) {
  await setBackgroundImage(fullUrl);
}

/**
 * Set the background image for the page
 */
async function setBackgroundImage(imageUrl) {
  try {
    await Word.run(async (context) => {
      const sections = context.document.sections;
      sections.load("items");
      await context.sync();

      sections.items.forEach((section) => {
        const background = section.getHeader("primary").paragraphs.getFirst();
        background.insertInlinePictureFromBase64(
          imageUrl,
          Word.InsertLocation.replace
        );
        // Style the background image (e.g., behind text)
        background.style = Word.StyleId.picture;
      });
    });
  } catch (error) {
    console.error("Error setting background image:", error);
  }
}

/**
 * Basic fetch -> blob
 */
async function fetchImageBlob(url) {
  const response = await fetch(url);
  return response.blob();
}

/**
 * Convert blob to base64
 */
function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => {
      const base64String = reader.result.split(",")[1];
      resolve(base64String);
    };
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}