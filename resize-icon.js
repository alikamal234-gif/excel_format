const Jimp = require('jimp');
const path = require('path');

const imagePath = path.join(__dirname, 'assets', 'icon.png');
const outputPath = path.join(__dirname, 'assets', 'icon_128x128.png');

async function resizeImage() {
  try {
    const image = await Jimp.read(imagePath);
    await image
      .resize(128, 128)
      .writeAsync(outputPath);
    console.log('Image successfully resized to 128x128 and saved to:', outputPath);
  } catch (err) {
    console.error('Error resizing image:', err);
  }
}

resizeImage();
