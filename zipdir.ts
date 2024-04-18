var zipdir = require("zip-dir");

zipdir("./temp", { saveTo: "./myzip.pptx" }, function (err, buffer) {
  console.log("doneso");
});
