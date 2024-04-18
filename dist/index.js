"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const officegen_1 = __importDefault(require("officegen"));
const fs_1 = __importDefault(require("fs"));
// Load an existing PowerPoint file
const inputFileName = "input.pptx";
const outputFileName = "output.pptx";
// Create a new PowerPoint object
const pptx = (0, officegen_1.default)("pptx");
// Create a new stream for the modified PowerPoint file
const outputStream = fs_1.default.createWriteStream(outputFileName);
// Read the existing PowerPoint file
const pptxReadStream = fs_1.default.createReadStream(inputFileName);
pptxReadStream.on("data", (chunk) => {
    // Process the chunk and modify the PowerPoint object as needed
    pptx.generate(outputStream, chunk);
});
pptxReadStream.on("end", () => {
    // Finalize the PowerPoint object and close the output stream
    pptx.generate(outputStream);
    outputStream.end();
});
pptx.on("finalize", () => {
    console.log("Modified PowerPoint file saved successfully.");
});
pptx.on("error", (err) => {
    console.log(err);
});
//# sourceMappingURL=index.js.map