const express = require("express");
const multer = require("multer");

const app = express();
const upload = multer({
  dest: "/tmp/uploads"
});

app.get("/", (req, res) => {
  res.json({
    success: true,
    service: "iLovePDF4 Free Conversion Engine",
    status: "online"
  });
});

app.post("/convert", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({
        success: false,
        error: "No file uploaded."
      });
    }

    return res.status(501).json({
      success: false,
      error: "Conversion engine is not connected yet."
    });

  } catch (error) {
    return res.status(500).json({
      success: false,
      error: error.message || "Conversion failed."
    });
  }
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log(`Conversion engine running on port ${PORT}`);
});
