var zipdir = require("zip-dir");
const fs = require("fs");
const yauzl = require("yauzl");
const xml2js = require("xml2js");

function savePptxAsFolder(pptxPath) {
  const tempDir = "repairedPPTX";

  // Step 1: Unzip the .pptx file
  yauzl.open(pptxPath, { lazyEntries: true }, (err, zipfile) => {
    if (err) throw err;

    zipfile.readEntry();
    zipfile.on("entry", (entry) => {
      const entryPath = entry.fileName;
      const destPath = `${tempDir}/${entryPath}`;

      if (/\/$/.test(entryPath)) {
        // Create directory if it doesn't exist
        fs.mkdirSync(destPath, { recursive: true });
        zipfile.readEntry();
      } else {
        // Extract file
        zipfile.openReadStream(entry, (err, readStream) => {
          if (err) throw err;
          readStream.pipe(fs.createWriteStream(destPath));
          readStream.on("end", () => zipfile.readEntry());
        });
      }
    });
  });
}

savePptxAsFolder("repaired.pptx");
