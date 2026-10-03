const express = require("express");
const multer = require("multer");
const fs = require("fs");
const path = require("path");
const { execFile } = require("child_process");

const {
  Document,
  Packer,
  Paragraph,
  ImageRun
} = require("docx");

const app = express();

app.use(express.json({ limit: "10mb" }));

app.use((req, res, next) => {
  res.header("Access-Control-Allow-Origin", "*");
  res.header(
    "Access-Control-Allow-Methods",
    "GET,POST,OPTIONS"
  );
  res.header(
    "Access-Control-Allow-Headers",
    "Content-Type"
  );

  if (req.method === "OPTIONS") {
    return res.sendStatus(204);
  }

  next();
});

const UPLOAD_DIR = "/tmp/uploads";
const OUTPUT_DIR = "/tmp/output";
const PROFILE_DIR = "/tmp/lo-profiles";

fs.mkdirSync(UPLOAD_DIR, { recursive: true });
fs.mkdirSync(OUTPUT_DIR, { recursive: true });
fs.mkdirSync(PROFILE_DIR, { recursive: true });

const upload = multer({
  dest: UPLOAD_DIR
});

const jobs = new Map();

const PUBLIC_BASE_URL =
  "https://free-conversion-engine.onrender.com";

/* =========================
   HOME
========================= */

app.get("/", (req, res) => {
  res.json({
    success: true,
    service: "iLovePDF4 Free Conversion Engine",
    status: "online"
  });
});

/* =========================
   STATUS
========================= */

app.get("/status/:jobId", (req, res) => {
  const jobId = req.params.jobId;
  const job = jobs.get(jobId);

  if (!job) {
    return res.status(404).json({
      status: "error",
      error: "Job not found."
    });
  }

  if (job.status === "finished") {
    return res.json({
      status: "finished",
      url:
        PUBLIC_BASE_URL +
        "/download/" +
        encodeURIComponent(jobId),
      filename: job.filename
    });
  }

  if (job.status === "error") {
    return res.json({
      status: "error",
      error:
        job.error ||
        "Conversion failed."
    });
  }

  return res.json({
    status: "processing"
  });
});

/* =========================
   DOWNLOAD
========================= */

app.get("/download/:jobId", (req, res) => {
  const jobId = req.params.jobId;
  const job = jobs.get(jobId);

  if (
    !job ||
    job.status !== "finished" ||
    !job.outputBuffer
  ) {
    return res
      .status(404)
      .send("Output file no longer exists.");
  }

  res.setHeader(
    "Content-Type",
    job.contentType ||
      "application/octet-stream"
  );

  res.setHeader(
    "Content-Disposition",
    'attachment; filename="' +
      job.filename.replace(/"/g, "") +
      '"'
  );

  res.setHeader(
    "Content-Length",
    job.outputBuffer.length
  );

  res.setHeader(
    "Cache-Control",
    "no-store, no-cache, must-revalidate"
  );

  res.setHeader(
    "Pragma",
    "no-cache"
  );

  res.send(job.outputBuffer);
});

/* =========================
   CONVERSION MAP
========================= */

const conversionMap = {
  "word-to-pdf": {
    extension: "pdf",
    contentType: "application/pdf"
  },

  "powerpoint-to-pdf": {
    extension: "pdf",
    contentType: "application/pdf"
  },

  "excel-to-pdf": {
    extension: "pdf",
    contentType: "application/pdf"
  },

  "pdf-to-word": {
    extension: "docx",
    contentType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  },

  "pdf-to-powerpoint": {
    extension: "pptx",
    contentType:
      "application/vnd.openxmlformats-officedocument.presentationml.presentation"
  },

  "pdf-to-excel": {
    extension: "xlsx",
    contentType:
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  }
};

/* =========================
   PDF -> WORD
   pdftoppm method
========================= */

async function convertScannedPdfToDocx(
  inputPath,
  outputPath
) {
  const os = require("os");

  const tempDir = fs.mkdtempSync(
    path.join(
      os.tmpdir(),
      "pdf-render-"
    )
  );

  try {
    const prefix =
      path.join(tempDir, "page");

    console.log(
      "Rendering PDF pages with pdftoppm..."
    );

    await new Promise(
      (resolve, reject) => {
        execFile(
          "pdftoppm",
          [
            "-png",
            "-r",
            "72",
            inputPath,
            prefix
          ],
          {
            timeout: 120000,
            maxBuffer:
              20 * 1024 * 1024
          },
          (
            error,
            stdout,
            stderr
          ) => {
            if (error) {
              reject(
                new Error(
                  stderr?.trim() ||
                    error.message ||
                    "PDF rendering failed."
                )
              );
              return;
            }

            resolve();
          }
        );
      }
    );

    const files = fs
      .readdirSync(tempDir)
      .filter(
        file =>
          /^page-\d+\.png$/i.test(
            file
          )
      )
      .sort((a, b) => {
        const na = parseInt(
          a.match(/(\d+)/)[1],
          10
        );

        const nb = parseInt(
          b.match(/(\d+)/)[1],
          10
        );

        return na - nb;
      });

    if (!files.length) {
      throw new Error(
        "No PDF pages were rendered."
      );
    }

    console.log(
      "Rendered pages:",
      files.length
    );

    function getPngSize(
      filePath
    ) {
      const buffer =
        fs.readFileSync(
          filePath
        );

      if (
        buffer.length < 24 ||
        buffer.readUInt32BE(0) !==
          0x89504e47
      ) {
        throw new Error(
          "Invalid PNG output."
        );
      }

      return {
        width:
          buffer.readUInt32BE(16),

        height:
          buffer.readUInt32BE(20)
      };
    }

    const sections = [];

    for (
      const file of files
    ) {
      const imagePath =
        path.join(
          tempDir,
          file
        );

      const imageBuffer =
        fs.readFileSync(
          imagePath
        );

      const {
        width,
        height
      } = getPngSize(
        imagePath
      );

      const pageWidthTwips =
        Math.round(
          width * 15
        );

      const pageHeightTwips =
        Math.round(
          height * 15
        );

      sections.push({
        properties: {
          page: {
            size: {
              width:
                pageWidthTwips,

              height:
                pageHeightTwips
            },

            margin: {
              top: 0,
              bottom: 0,
              left: 0,
              right: 0,
              header: 0,
              footer: 0,
              gutter: 0
            }
          }
        },

        children: [
          new Paragraph({
            spacing: {
              before: 0,
              after: 0,
              line: 0
            },

            indent: {
              left: 0,
              right: 0,
              firstLine: 0
            },

            children: [
              new ImageRun({
                type: "png",
                data: imageBuffer,
                transformation: {
                  width,
                  height
                }
              })
            ]
          })
        ]
      });
    }

    const doc =
      new Document({
        sections
      });

    const buffer =
      await Packer.toBuffer(
        doc
      );

    if (
      !buffer ||
      buffer.length < 1000
    ) {
      throw new Error(
        "DOCX output is invalid or empty."
      );
    }

    fs.writeFileSync(
      outputPath,
      buffer
    );

    console.log(
      "PDF -> WORD completed:",
      outputPath
    );

  } finally {
    cleanupDirectory(
      tempDir
    );
  }
}

/* =========================
   MAIN CONVERT
========================= */

app.post(
  "/convert",
  upload.single("file"),
  (req, res) => {

    if (!req.file) {
      return res.status(400).json({
        success: false,
        error:
          "No file uploaded."
      });
    }

    const tool =
      String(
        req.body.tool || ""
      )
        .trim()
        .toLowerCase();

    const jobId =
      String(
        req.body.jobId || ""
      ).trim();

    if (!jobId) {
      safeDelete(
        req.file.path
      );

      return res.status(400).json({
        success: false,
        error:
          "Missing job id."
      });
    }

    const conversion =
      conversionMap[tool];

    if (!conversion) {
      safeDelete(
        req.file.path
      );

      return res.status(400).json({
        success: false,
        error:
          "Unsupported conversion tool: " +
          tool
      });
    }

    const originalName =
      String(
        req.body.filename ||
          req.file.originalname ||
          "file"
      );

    const safeName =
      path.basename(
        originalName
      );

    const inputPath =
      path.join(
        UPLOAD_DIR,
        jobId +
          "-" +
          safeName
      );

    const jobOutputDir =
      path.join(
        OUTPUT_DIR,
        jobId
      );

    const jobProfileDir =
      path.join(
        PROFILE_DIR,
        jobId
      );

    fs.mkdirSync(
      jobOutputDir,
      {
        recursive: true
      }
    );

    fs.mkdirSync(
      jobProfileDir,
      {
        recursive: true
      }
    );

    try {
      fs.renameSync(
        req.file.path,
        inputPath
      );
    } catch (error) {

      safeDelete(
        req.file.path
      );

      return res.status(500).json({
        success: false,
        error:
          "Unable to prepare uploaded file."
      });
    }

    const outputFilename =
      path.parse(
        safeName
      ).name +
      "." +
      conversion.extension;

    jobs.set(
      jobId,
      {
        status:
          "processing",

        filename:
          outputFilename
      }
    );

    res.json({
      success: true,
      jobId: jobId,
      status:
        "processing"
    });

    let finished = false;

    /* =========================
       SUCCESS
    ========================= */

    function finishSuccess(
      outputFile
    ) {
      if (finished) {
        return;
      }

      try {
        if (
          !fs.existsSync(
            outputFile
          )
        ) {
          throw new Error(
            "Output file was not created."
          );
        }

        const stat =
          fs.statSync(
            outputFile
          );

        if (
          !stat.isFile() ||
          stat.size < 1000
        ) {
          throw new Error(
            "Output file is incomplete or empty."
          );
        }

        const outputBuffer =
          fs.readFileSync(
            outputFile
          );

        if (
          !outputBuffer ||
          outputBuffer.length < 1000
        ) {
          throw new Error(
            "Output file is invalid."
          );
        }

        finished = true;

        jobs.set(
          jobId,
          {
            status:
              "finished",

            outputBuffer:
              outputBuffer,

            filename:
              outputFilename,

            contentType:
              conversion.contentType
          }
        );

        console.log(
          "CONVERSION FINISHED:",
          jobId
        );

        console.log(
          "Output size:",
          outputBuffer.length,
          "bytes"
        );

      } catch (error) {

        finished = true;

        jobs.set(
          jobId,
          {
            status:
              "error",

            error:
              error.message ||
              "Unable to read conversion output."
          }
        );

        console.error(
          "Output validation error:",
          error.message
        );
      }

      safeDelete(
        inputPath
      );

      cleanupDirectory(
        jobOutputDir
      );

      cleanupDirectory(
        jobProfileDir
      );
    }

    /* =========================
       ERROR
    ========================= */

    function finishError(
      message
    ) {
      if (finished) {
        return;
      }

      finished = true;

      jobs.set(
        jobId,
        {
          status:
            "error",

          error:
            message
        }
      );

      safeDelete(
        inputPath
      );

      cleanupDirectory(
        jobOutputDir
      );

      cleanupDirectory(
        jobProfileDir
      );

      console.error(
        "CONVERSION ERROR:",
        message
      );
    }

    /* =========================
       PDF -> WORD
========================= */

    if (
      tool === "pdf-to-word"
    ) {

      const outputFile =
        path.join(
          jobOutputDir,
          outputFilename
        );

      console.log(
        "================================="
      );

      console.log(
        "Starting PDF -> WORD conversion"
      );

      console.log(
        "Job:",
        jobId
      );

      console.log(
        "Input:",
        inputPath
      );

      console.log(
        "Output:",
        outputFile
      );

      console.log(
        "================================="
      );

      convertScannedPdfToDocx(
        inputPath,
        outputFile
      )
        .then(() => {

          if (finished) {
            return;
          }

          finishSuccess(
            outputFile
          );
        })

        .catch(error => {

          console.error(
            "PDF TO WORD ERROR:",
            error
          );

          finishError(
            error?.message ||
            "PDF to Word conversion failed."
          );
        });

      return;
    }

    /* =========================
       LIBREOFFICE CONVERSIONS
========================= */

    let libreOfficeArgs = [
  "--headless",
  "--nologo",
  "--nodefault",
  "--nofirststartwizard",
  "--nolockcheck",
  "-env:UserInstallation=file://" + jobProfileDir,
  "--convert-to"
];

if (tool === "word-to-pdf") {
  libreOfficeArgs.push(
    "pdf:writer_pdf_Export"
  );
} else if (tool === "powerpoint-to-pdf") {
  libreOfficeArgs.push(
    "pdf:impress_pdf_Export"
  );
} else if (tool === "excel-to-pdf") {
  libreOfficeArgs.push(
    "pdf:calc_pdf_Export"
  );
} else if (tool === "pdf-to-powerpoint") {
  libreOfficeArgs.push(
    "pptx:Impress MS PowerPoint 2007 XML"
  );

  libreOfficeArgs.push(
    "--infilter=draw_pdf_import"
  );
} else if (tool === "pdf-to-excel") {
  libreOfficeArgs.push(
    "xlsx:Calc MS Excel 2007 XML"
  );

  libreOfficeArgs.push(
    "--infilter=draw_pdf_import"
  );
} else {
  finishError(
    "Unsupported conversion tool."
  );
  return;
}

libreOfficeArgs.push(
  "--outdir",
  jobOutputDir,
  inputPath
);

console.log("=================================");
console.log("Starting conversion:", tool);
console.log("Job:", jobId);
console.log("Input:", inputPath);
console.log("Output directory:", jobOutputDir);
console.log("=================================");

const child = execFile(
  "libreoffice",
  libreOfficeArgs,
  {
    timeout: 120000,
    maxBuffer: 50 * 1024 * 1024
  },
  (error, stdout, stderr) => {
    console.log(
      "LibreOffice stdout:",
      stdout || ""
    );

    console.log(
      "LibreOffice stderr:",
      stderr || ""
    );

    if (finished) {
      return;
    }

    const outputFile = findOutputFile(
      jobOutputDir,
      conversion.extension
    );

    if (outputFile) {
      try {
        const stat = fs.statSync(outputFile);

        console.log(
          "Output file found:",
          outputFile
        );

        console.log(
          "Output size:",
          stat.size,
          "bytes"
        );

        if (
          stat.isFile() &&
          stat.size >= 1000
        ) {
          finishSuccess(outputFile);
          return;
        }
      } catch (e) {
        console.error(
          "Output check failed:",
          e.message
        );
      }
    }

    if (error) {
      finishError(
        getLibreOfficeError(
          tool,
          stderr,
          error
        )
      );
      return;
    }

    finishError(
      "LibreOffice did not create a valid output file."
    );
  }
);

setTimeout(() => {
  if (finished) {
    return;
  }

  try {
    if (child && !child.killed) {
      child.kill("SIGKILL");
    }
  } catch {}

  finishError(
    "LibreOffice conversion timed out."
  );
}, 120000);

setTimeout(() => {
  jobs.delete(jobId);
}, 10 * 60 * 1000);
      }

);

/* =========================
   FIND OUTPUT FILE
========================= */

function findOutputFile(
  directory,
  extension
) {
  try {

    if (
      !fs.existsSync(
        directory
      )
    ) {
      return null;
    }

    const files =
      fs.readdirSync(
        directory
      );

    const wanted =
      "." +
      extension.toLowerCase();

    const validFiles =
      files.filter(
        file =>
          path
            .extname(file)
            .toLowerCase() ===
          wanted
      );

    if (!validFiles.length) {
      return null;
    }

    /*
     * Choose the newest valid file.
     */

    validFiles.sort(
      (a, b) => {

        const statA =
          fs.statSync(
            path.join(
              directory,
              a
            )
          );

        const statB =
          fs.statSync(
            path.join(
              directory,
              b
            )
          );

        return (
          statB.mtimeMs -
          statA.mtimeMs
        );
      }
    );

    return path.join(
      directory,
      validFiles[0]
    );

  } catch {
    return null;
  }
}

/* =========================
   SAFE DELETE
========================= */

function safeDelete(
  filePath
) {
  try {

    if (
      filePath &&
      fs.existsSync(
        filePath
      )
    ) {
      fs.unlinkSync(
        filePath
      );
    }

  } catch {}
}

/* =========================
   CLEAN DIRECTORY
========================= */

function cleanupDirectory(
  directory
) {
  try {

    if (
      !fs.existsSync(
        directory
      )
    ) {
      return;
    }

    fs.rmSync(
      directory,
      {
        recursive: true,
        force: true
      }
    );

  } catch {}
}

/* =========================
   LIBREOFFICE ERROR
========================= */

function getLibreOfficeError(
  tool,
  stderr,
  error
) {

  if (
    stderr &&
    stderr.trim()
  ) {
    return (
      "LibreOffice conversion failed: " +
      stderr.trim()
    );
  }

  return (
    "Conversion failed for " +
    tool +
    ". " +
    (
      error?.message ||
      "Unknown conversion error."
    )
  );
}

/* =========================
   SERVER
========================= */

const PORT =
  process.env.PORT ||
  10000;

app.listen(
  PORT,
  "0.0.0.0",
  () => {

    console.log(
      "================================="
    );

    console.log(
      "iLovePDF4 Free Conversion Engine"
    );

    console.log(
      "Running on port " +
        PORT
    );

    console.log(
      "================================="
    );
  }
);
