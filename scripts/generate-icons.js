/**
 * Writes manifest icon PNGs into /assets (solid brand blue).
 * Run before webpack so Office can load IconUrl / ribbon images from the deployed host.
 */
const fs = require("fs");
const path = require("path");
const { PNG } = require("pngjs");

const assetsDir = path.join(__dirname, "..", "assets");
const sizes = [16, 32, 64, 80];

fs.mkdirSync(assetsDir, { recursive: true });

for (const size of sizes) {
  const png = new PNG({ width: size, height: size });
  for (let y = 0; y < size; y++) {
    for (let x = 0; x < size; x++) {
      const idx = (size * y + x) << 2;
      png.data[idx] = 0x25;
      png.data[idx + 1] = 0x63;
      png.data[idx + 2] = 0xeb;
      png.data[idx + 3] = 255;
    }
  }
  fs.writeFileSync(path.join(assetsDir, `icon-${size}.png`), PNG.sync.write(png));
}
