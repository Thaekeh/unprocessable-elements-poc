const fs = require("fs");
const yauzl = require("yauzl");
const xml2js = require("xml2js");
var zipdir = require("zip-dir");

const removeElements: boolean = true;

function removeTextElement(pptxPath, slideNumber) {
  const tempDir = "temp";

  // Unzip the .pptx file
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

    zipfile.on("end", () => {
      // Open the XML file of the target slide
      const slideXmlPath = `${tempDir}/ppt/slides/slide${slideNumber}.xml`;
      const slideRelsPath = `${tempDir}/ppt/slides/_rels/slide${slideNumber}.xml.rels`;
      fs.readFile(slideXmlPath, "utf8", (err, data) => {
        if (err) throw err;

        // Parse XML
        xml2js.parseString(data, (err, result) => {
          if (err) throw err;

          if (removeElements) {
            result["p:sld"]["p:cSld"][0]["p:spTree"][0]["p:sp"].forEach(
              (shape, i) => {
                result["p:sld"]["p:cSld"][0]["p:spTree"][0]["p:sp"][i][
                  "p:txBody"
                ] = [];
              }
            );

            const path = findPath(result, "p:pic");

            const images = result["p:sld"]["p:cSld"][0]["p:spTree"][0]["p:pic"];

            if (images) {
              console.log(
                `pic`,
                result["p:sld"]["p:cSld"][0]["p:spTree"][0]["p:pic"][0][
                  "p:blipFill"
                ]
              );

              delete result["p:sld"]["p:cSld"][0]["p:spTree"][0]["p:pic"][0][
                "p:blipFill"
              ];
              delete result["p:sld"]["p:cSld"][0]["p:spTree"][0]["p:pic"][0][
                "p:nvPicPr"
              ];
              delete result["p:sld"]["p:cSld"][0]["p:spTree"][0]["p:pic"][0][
                "p:spPr"
              ];

              console.log(
                `pic`,
                result["p:sld"]["p:cSld"][0]["p:spTree"][0]["p:pic"]
              );

              fs.readFile(slideRelsPath, "utf8", (err, data) => {
                xml2js.parseString(data, (err, newResult) => {
                  if (newResult) {
                    const relationship = newResult["Relationships"][
                      "Relationship"
                    ].find((relationship) => relationship.$.Id === "rId3");
                    console.log("relationship", relationship);
                    // remove relationship from newResults
                    //   console.log(
                    //     "relations",
                    //     newResult["Relationships"]["Relationship"]
                    //   );
                    const rels = newResult["Relationships"][
                      "Relationship"
                    ].filter(
                      (relationship) => !relationship.$.Target.includes("media")
                    );

                    newResult["Relationships"]["Relationship"] = rels;

                    //   console.log("rels", rels);

                    const builder = new xml2js.Builder();
                    const modifiedXml = builder.buildObject(newResult);
                    fs.writeFile(slideRelsPath, modifiedXml, "utf8", (err) => {
                      console.log("error", err);
                    });
                  } else {
                    console.log("no rels");
                  }
                });
              });
            }
          }

          // Save the modified XML file
          const builder = new xml2js.Builder();
          const modifiedXml = builder.buildObject(result);
          fs.writeFile(slideXmlPath, modifiedXml, "utf8", (err) => {
            if (err) throw err;

            // Repackage the .pptx file
            // (you can use a library like `zip-dir` to zip the contents back)
            console.log("Text element removed successfully.");

            zipdir(
              "./temp",
              { saveTo: "./output.pptx" },
              function (err, buffer) {
                console.log("doneso");
              }
            );
          });
        });
      });
    });
  });
}

for (let i = 1; i < 6; i++) {
  removeTextElement("input.pptx", i);
}
// removeTextElement("input.pptx", 1, "Lesson 1");

const findPath = (obj, targetTag, currentPath = []): string => {
  if (obj instanceof Object) {
    for (const key in obj) {
      if (key === targetTag) {
        // Found the target tag, return the current path
        const finalPath = "[" + currentPath.join("][") + "]";
        console.log("Path to the first <p:pic> tag:", finalPath);
        return finalPath;
      } else {
        // Continue searching recursively
        findPath(obj[key], targetTag, [...currentPath, key]);
      }
    }
  }
};
