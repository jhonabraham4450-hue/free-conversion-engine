const express = require("express");
const multer = require("multer");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
const os = require("os");
const { execFile } = require("child_process");
const pptxgen = require("pptxgenjs");

const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  ImageRun
} = require("docx");

const app = express();

app.use(cors());

app.use(
  express.json({
    limit: "50mb"
  })
);

app.use(
  express.urlencoded({
    extended: true,
    limit: "50mb"
  })
);


// ==================================================
// DIRECTORIES
// ==================================================

const uploadDir =
  path.join(
    os.tmpdir(),
    "uploads"
  );

const outputDir =
  path.join(
    os.tmpdir(),
    "output"
  );

fs.mkdirSync(
  uploadDir,
  {
    recursive: true
  }
);

fs.mkdirSync(
  outputDir,
  {
    recursive: true
  }
);


// ==================================================
// MULTER
// ==================================================

const upload =
  multer({
    dest: uploadDir,

    limits: {
      fileSize:
        512 * 1024 * 1024
    }
  });


// ==================================================
// JOB STORAGE
// ==================================================

const jobs =
  new Map();


// ==================================================
// ROOT
// ==================================================

app.get(
  "/",
  (req, res) => {

    res.json({
      success: true,
      service:
        "iLovePDF4 Free Conversion Engine",
      status: "online"
    });

  }
);


// ==================================================
// STATUS
// ==================================================

app.get(
  "/status/:jobId",
  (req, res) => {

    const job =
      jobs.get(
        req.params.jobId
      );

    if (!job) {

      return res.status(404).json({
        success: false,
        error:
          "Job not found."
      });

    }

    res.json({
      success: true,

      jobId:
        req.params.jobId,

      status:
        job.status,

      error:
        job.error || null,

      filename:
        job.filename || null
    });

  }
);


// ==================================================
// DOWNLOAD
// ==================================================

app.get(
  "/download/:jobId",
  (req, res) => {

    const job =
      jobs.get(
        req.params.jobId
      );

    if (!job) {

      return res.status(404).json({
        success: false,
        error:
          "Job not found."
      });

    }

    if (
      job.status !==
      "finished"
    ) {

      return res.status(400).json({
        success: false,
        error:
          "Conversion is not finished yet."
      });

    }

    if (
      !job.outputPath ||
      !fs.existsSync(
        job.outputPath
      )
    ) {

      return res.status(404).json({
        success: false,
        error:
          "Output file not found."
      });

    }

    res.download(
      job.outputPath,

      job.filename ||
        "converted-file",

      err => {

        if (err) {

          console.error(
            "Download error:",
            err.message
          );

        }

      }
    );

  }
);


// ==================================================
// CONVERSION MAP
// ==================================================

const conversionMap = {

  "word-to-pdf": {

    input: [
      "doc",
      "docx"
    ],

    output: "pdf",

    extension:
      ".pdf"

  },


  "powerpoint-to-pdf": {

    input: [
      "ppt",
      "pptx"
    ],

    output: "pdf",

    extension:
      ".pdf"

  },


  "excel-to-pdf": {

    input: [
      "xls",
      "xlsx"
    ],

    output: "pdf",

    extension:
      ".pdf"

  },


  "pdf-to-word": {

    input: [
      "pdf"
    ],

    output: "docx",

    extension:
      ".docx"

  },


  "pdf-to-powerpoint": {

    input: [
      "pdf"
    ],

    output: "pptx",

    extension:
      ".pptx"

  },


  "pdf-to-excel": {

    input: ["pdf"],
    output: "xlsx",
    extension: ".xlsx"

  },

  "protect-pdf": {
    input:["pdf"],
    output:"pdf",
    extension:".pdf"
  },

  "unlock-pdf": {
    input:["pdf"],
    output:"pdf",
    extension:".pdf"
  },

  "ocr-pdf": {
    input:["pdf"],
    output:"pdf",
    extension:".pdf"
  }

};


// ==================================================
// CLEANUP DIRECTORY
// ==================================================

function cleanupDirectory(
  dir
) {

  try {

    if (
      !fs.existsSync(dir)
    ) {

      return;

    }

    const files =
      fs.readdirSync(dir);

    for (
      const file of files
    ) {

      const filePath =
        path.join(
          dir,
          file
        );

      try {

        const stat =
          fs.statSync(
            filePath
          );

        if (
          stat.isDirectory()
        ) {

          cleanupDirectory(
            filePath
          );

          fs.rmdirSync(
            filePath
          );

        } else {

          fs.unlinkSync(
            filePath
          );

        }

      } catch {}

    }

    try {

      fs.rmdirSync(
        dir
      );

    } catch {}

  } catch {}

}


// ==================================================
// FIND OUTPUT FILE
// ==================================================

function findOutputFile(
  dir,
  extension
) {

  try {

    if (
      !fs.existsSync(dir)
    ) {

      return null;

    }

    const files =
      fs.readdirSync(dir);

    const wantedExtension =
      extension.toLowerCase();

    for (
      const file of files
    ) {

      const fullPath =
        path.join(
          dir,
          file
        );

      try {

        const stat =
          fs.statSync(
            fullPath
          );

        if (
          stat.isFile() &&
          path.extname(file)
            .toLowerCase() ===
            wantedExtension
        ) {

          return fullPath;

        }

      } catch {}

    }

  } catch {}

  return null;

}


// ==================================================
// PDF -> WORD
// FAST TEXT FIRST
// OCR FALLBACK SECOND
// ==================================================

async function convertScannedPdfToDocx(
  inputPath,
  outputPath
) {

  // ==================================================
  // FAST TEXT EXTRACTION
  // ==================================================

  try {

    const textResult =
      await new Promise(
        (resolve, reject) => {

          execFile(
            "pdftotext",

            [
              "-layout",
              inputPath,
              "-"
            ],

            {
              timeout:
                15000,

              maxBuffer:
                50 * 1024 * 1024
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
                    "pdftotext failed."
                  )
                );

                return;

              }

              resolve(
                stdout || ""
              );

            }
          );

        }
      );


    const extractedText =
      String(
        textResult || ""
      )
      .replace(
        /\r/g,
        ""
      )
      .trim();


    // ==================================================
    // TEXT PDF
    // ==================================================

    if (
      extractedText
    ) {

      console.log(
        "PDF -> WORD: selectable text detected."
      );

      console.log(
        "Using fast text conversion."
      );


      const pages =
        extractedText.split(
          "\f"
        );


      const sections =
        [];


      for (
        let pageIndex = 0;
        pageIndex < pages.length;
        pageIndex++
      ) {

        const pageText =
          pages[
            pageIndex
          ].trim();


        const lines =
          pageText
            ? pageText.split(
                "\n"
              )
            : [""];


        const children =
          [];


        for (
          const line of lines
        ) {

          children.push(

            new Paragraph({

              spacing: {

                before: 0,

                after: 0

              },

              children: [

                new TextRun({

                  text:
                    line

                })

              ]

            })

          );

        }


        sections.push({

          children

        });

      }


      const doc =
        new Document({

          sections:
            sections.length
              ? sections
              : [

                  {

                    children: [

                      new Paragraph({

                        children: [

                          new TextRun({

                            text:
                              extractedText

                          })

                        ]

                      })

                    ]

                  }

                ]

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
        "PDF -> WORD fast conversion completed:",
        outputPath
      );


      console.log(
        "DOCX size:",
        buffer.length,
        "bytes"
      );


      return;

    }


    console.log(
      "PDF -> WORD: no selectable text found."
    );

    console.log(
      "Switching to OCR."
    );


  } catch (error) {

    console.log(
      "Fast PDF text extraction unavailable."
    );

    console.log(
      "Switching to OCR:",
      error.message
    );

  }


  // ==================================================
  // OCR FALLBACK
  // FOR SCANNED PDF
  // ==================================================

  const tempDir =
    fs.mkdtempSync(
      path.join(
        os.tmpdir(),
        "pdf-render-"
      )
    );


  try {

    const prefix =
      path.join(
        tempDir,
        "page"
      );


    console.log(
      "================================="
    );

    console.log(
      "PDF -> WORD + OCR"
    );

    console.log(
      "Input:",
      inputPath
    );

    console.log(
      "================================="
    );


    // ==================================================
    // RENDER PDF - PDF -> WORD ONLY
    // ==================================================

    async function renderPdf(command) {

      return new Promise(
        (
          resolve,
          reject
        ) => {

          execFile(
            command,

            [
              "-png",
              "-r",
              "72",
              inputPath,
              prefix
            ],

            {
              timeout:
                180000,

              maxBuffer:
                50 * 1024 * 1024
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
                    `${command} rendering failed.`
                  )
                );

                return;

              }

              resolve();

            }
          );

        }
      );

    }


    try {

      await renderPdf(
        "pdftocairo"
      );

      console.log(
        "PDF rendered using pdftocairo."
      );

    } catch (
      firstError
    ) {

      console.log(
        "pdftocairo failed. Trying pdftoppm..."
      );

      try {

        await renderPdf(
          "pdftoppm"
        );

        console.log(
          "PDF rendered using pdftoppm."
        );

      } catch (
        secondError
      ) {

        throw new Error(
          "PDF rendering failed. Both pdftocairo and pdftoppm failed."
        );

      }

    }


    // ==================================================
    // FIND PAGES
    // ==================================================

    const files =
      fs
        .readdirSync(
          tempDir
        )
        .filter(
          file =>
            /^page-\d+\.png$/i.test(
              file
            )
        )
        .sort(
          (
            a,
            b
          ) => {

            const na =
              parseInt(
                a.match(
                  /(\d+)/
                )[1],
                10
              );


            const nb =
              parseInt(
                b.match(
                  /(\d+)/
                )[1],
                10
              );


            return na - nb;

          }
        );


    if (
      !files.length
    ) {

      throw new Error(
        "No PDF pages were rendered."
      );

    }


    console.log(
      "Rendered pages:",
      files.length
    );


    // ==================================================
    // PNG SIZE
    // ==================================================

    function getPngSize(
      filePath
    ) {

      const buffer =
        fs.readFileSync(
          filePath
        );


      if (
        buffer.length < 24 ||
        buffer.readUInt32BE(
          0
        ) !==
        0x89504e47
      ) {

        throw new Error(
          "Invalid PNG output."
        );

      }


      return {

        width:
          buffer.readUInt32BE(
            16
          ),

        height:
          buffer.readUInt32BE(
            20
          )

      };

    }


    // ==================================================
    // OCR
    // ==================================================

    async function runOCR(
      imagePath
    ) {

      return new Promise(
        resolve => {

          execFile(

            "tesseract",

            [
              imagePath,

              "stdout",

              "-l",

              "ben+eng",

              "--psm",

              "6"
            ],

            {

              timeout:
                20000,

              maxBuffer:
                20 * 1024 * 1024

            },

            (
              error,
              stdout,
              stderr
            ) => {

              if (error) {

                console.error(
                  "OCR failed:",
                  stderr ||
                  error.message
                );

                resolve(
                  ""
                );

                return;

              }


              resolve(

                (
                  stdout ||
                  ""
                )
                .replace(
                  /\r/g,
                  ""
                )
                .trim()

              );

            }
          );

        }
      );

    }
        // ==================================================
    // CREATE DOCX SECTIONS
    // ==================================================

    const sections =
      [];



    for (
      let i = 0;
      i < files.length;
      i++
    ) {

      const file =
        files[i];


      const imagePath =
        path.join(
          tempDir,
          file
        );


      console.log(
        `Processing page ${i + 1}/${files.length}`
      );


      const jpgPath =
        path.join(
          tempDir,
          `page-${i + 1}.jpg`
        );


      await new Promise(
        (
          resolve,
          reject
        ) => {

          execFile(
            "convert",

            [
              imagePath,
              "-quality",
              "82",
              "-strip",
              jpgPath
            ],

            {
              timeout:
                60000,

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
                    "Image compression failed."
                  )
                );

                return;

              }

              resolve();

            }
          );

        }
      );


      const imageBuffer =
        fs.readFileSync(
          jpgPath
        );


      const {
        width,
        height
      } =
        getPngSize(
          imagePath
        );


      console.log(
        `Running OCR on page ${i + 1}...`
      );


      const ocrText =
        await runOCR(
          imagePath
        );


      console.log(
        `OCR page ${i + 1}:`,
        ocrText
          ? `${ocrText.length} characters`
          : "No text detected"
      );


      const pageWidthTwips =
        Math.round(
          width * 15
        );


      const pageHeightTwips =
        Math.round(
          height * 15
        );


      const children =
        [];


      // ==================================================
      // PAGE IMAGE
      // ==================================================

      children.push(

        new Paragraph({

          spacing: {

            before: 0,

            after: 0

          },

          children: [

            new ImageRun({

              type:
                "png",

              data:
                imageBuffer,

              transformation: {

                width:
                  Math.max(
                    100,
                    width - 8
                  ),

                height:
                  Math.max(
                    100,
                    height - 8
                  )

              }

            })

          ]

        })

      );


      // ==================================================
      // OCR TEXT
      // ==================================================

      if (
        ocrText
      ) {

        children.push(

          new Paragraph({

            spacing: {

              before:
                100,

              after:
                100

            },

            children: [

              new TextRun({

                text:
                  ocrText,

                size:
                  20

              })

            ]

          })

        );

      }


      // ==================================================
      // SECTION
      // ==================================================

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

              top:
                120,

              bottom:
                120,

              left:
                120,

              right:
                120,

              header:
                0,

              footer:
                0,

              gutter:
                0

            }

          }

        },

        children

      });

    }


    // ==================================================
    // CREATE DOCX
    // ==================================================

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
      "PDF -> WORD + OCR completed:",
      outputPath
    );


    console.log(
      "DOCX size:",
      buffer.length,
      "bytes"
    );


  } finally {

    cleanupDirectory(
      tempDir
    );

  }

}


// ==================================================
// LIBREOFFICE ERROR
// ==================================================

function getLibreOfficeError(
  tool,
  stderr,
  error
) {

  const message =
    stderr?.trim() ||
    error?.message ||
    "LibreOffice conversion failed.";


  return (
    `${tool} conversion failed: ${message}`
  );

}


// ==================================================
// MAIN CONVERT API
// ==================================================

app.post(
  "/convert",

  upload.single("file"),

  async (
    req,
    res
  ) => {

    let inputPath =
      null;


    try {

      // ------------------------------------------------
      // CHECK FILE
      // ------------------------------------------------

      if (
        !req.file
      ) {

        return res.status(400).json({

          success:
            false,

          error:
            "No file uploaded."

        });

      }


      inputPath =
        req.file.path;


      // ------------------------------------------------
      // PRESERVE ORIGINAL EXTENSION
      // ------------------------------------------------

      const originalExt =
        path.extname(
          req.file.originalname || ""
        ).toLowerCase();


      if (
        originalExt
      ) {

        const renamedInputPath =
          inputPath +
          originalExt;


        fs.renameSync(
          inputPath,
          renamedInputPath
        );


        inputPath =
          renamedInputPath;

      }


      // ------------------------------------------------
      // TOOL
      // ------------------------------------------------

      const tool =
        String(
          req.body.tool ||
          req.query.tool ||
          ""
        )
        .trim()
        .toLowerCase();


      console.log(
        "================================="
      );


      console.log(
        "Starting conversion:",
        tool
      );


      console.log(
        "Original filename:",
        req.file.originalname
      );


      console.log(
        "Input:",
        inputPath
      );


      console.log(
        "================================="
      );


      // ------------------------------------------------
      // CHECK TOOL
      // ------------------------------------------------

      const conversion =
        conversionMap[
          tool
        ];


      if (
        !conversion
      ) {

        try {

          fs.unlinkSync(
            inputPath
          );

        } catch {}


        return res.status(400).json({

          success:
            false,

          error:
            "Unsupported conversion tool."

        });

      }


      // ------------------------------------------------
      // CREATE JOB
      // ------------------------------------------------

      const jobId =
        `${Date.now()}-${Math.random()
          .toString(36)
          .slice(2, 10)}`;


      const jobDir =
        path.join(
          outputDir,
          jobId
        );


      fs.mkdirSync(
        jobDir,
        {
          recursive:
            true
        }
      );


      const jobOutputDir =
        jobDir;


      // ------------------------------------------------
      // OUTPUT NAME
      // ------------------------------------------------

      const originalName =
        path.parse(
          req.file.originalname ||
          "converted-file"
        ).name;


      const safeName =
        originalName.replace(
          /[^a-zA-Z0-9._-]/g,
          "_"
        );


      const outputPath =
        path.join(
          jobOutputDir,
          `${safeName}${conversion.extension}`
        );


      const outputFilename =
        `${safeName}${conversion.extension}`;


      jobs.set(
        jobId,
        {

          status:
            "processing",

          outputPath:
            null,

          filename:
            outputFilename,

          error:
            null

        }
      );


      // ------------------------------------------------
      // RESPONSE
      // ------------------------------------------------

      res.json({

        success:
          true,

        jobId,

        status:
          "processing",

        filename:
          outputFilename

      });


      // ==================================================
      // PDF -> WORD
      // ==================================================

      if (
        tool ===
        "pdf-to-word"
      ) {

        try {

          await convertScannedPdfToDocx(
            inputPath,
            outputPath
          );


          if (
            !fs.existsSync(
              outputPath
            )
          ) {

            throw new Error(
              "PDF to Word output was not created."
            );

          }


          const stat =
            fs.statSync(
              outputPath
            );


          if (
            !stat.isFile() ||
            stat.size < 1000
          ) {

            throw new Error(
              "PDF to Word output is invalid."
            );

          }


          jobs.set(
            jobId,
            {

              status:
                "finished",

              outputPath:
                outputPath,

              filename:
                outputFilename,

              error:
                null

            }
          );


          console.log(
            "PDF -> WORD FINISHED:",
            outputPath
          );


          console.log(
            "DOCX size:",
            stat.size,
            "bytes"
          );


        } catch (
          error
        ) {

          console.error(
            "PDF -> WORD error:",
            error
          );


          jobs.set(
            jobId,
            {

              status:
                "error",

              outputPath:
                null,

              filename:
                outputFilename,

              error:
                error.message ||
                "PDF to Word conversion failed."

            }
          );

        }


        try {

          fs.unlinkSync(
            inputPath
          );

        } catch {}


        setTimeout(
          () => {

            jobs.delete(
              jobId
            );

            cleanupDirectory(
              jobDir
            );

          },

          10 * 60 * 1000

        );


        return;

      }
            // ==================================================
      // EXTRA PDF TOOLS
      // protect / unlock / OCR
      // ==================================================

      if (["protect-pdf","unlock-pdf","ocr-pdf"].includes(tool)) {

        try {

          const password =
            String(
              req.body.password || ""
            );


          if (
            tool === "protect-pdf" &&
            !password
          ) {

            throw new Error(
              "Password is required for Protect PDF."
            );

          }


          const run =
            (
              command,
              args,
              timeout = 300000
            ) =>
              new Promise(
                (
                  resolve,
                  reject
                ) => {

                  execFile(
                    command,
                    args,
                    {
                      timeout,
                      maxBuffer:
                        50 * 1024 * 1024
                    },

                    (
                      error,
                      stdout,
                      stderr
                    ) => {

                      if (error) {

                        reject(
                          new Error(
                            (
                              stderr ||
                              error.message ||
                              `${command} failed`
                            ).trim()
                          )
                        );

                        return;

                      }

                      resolve(
                        stdout
                      );

                    }
                  );

                }
              );


          // ------------------------------------------------
          // PROTECT PDF
          // ------------------------------------------------

          if (
            tool === "protect-pdf"
          ) {

            await run(
              "qpdf",

              [
                "--encrypt",
                "",
                password,
                "256",
                "--",
                inputPath,
                outputPath
              ]
            );

          }


          // ------------------------------------------------
          // UNLOCK PDF
          // ------------------------------------------------

          if (
            tool === "unlock-pdf"
          ) {

            if (
              password
            ) {

              await run(
                "qpdf",

                [
                  "--password=" +
                    password,

                  "--decrypt",

                  inputPath,

                  outputPath
                ]
              );

            } else {

              await run(
                "qpdf",

                [
                  "--decrypt",

                  inputPath,

                  outputPath
                ]
              );

            }

          }


          // ------------------------------------------------
          // OCR PDF
          // ------------------------------------------------

          if (
            tool === "ocr-pdf"
          ) {

            await run(
              "ocrmypdf",

              [
                "--force-ocr",

                "--deskew",

                "--optimize",
                "1",

                inputPath,

                outputPath
              ],

              600000
            );

          }


          if (
            !fs.existsSync(
              outputPath
            ) ||
            fs.statSync(
              outputPath
            ).size < 100
          ) {

            throw new Error(
              "Output PDF was not created correctly."
            );

          }


          jobs.set(
            jobId,
            {

              status:
                "finished",

              outputPath,

              filename:
                outputFilename,

              error:
                null

            }
          );


        } catch (
          error
        ) {

          console.error(
            `${tool} error:`,
            error
          );


          jobs.set(
            jobId,
            {

              status:
                "error",

              outputPath:
                null,

              filename:
                outputFilename,

              error:
                error.message ||
                `${tool} failed.`

            }
          );

        }


        try {

          fs.unlinkSync(
            inputPath
          );

        } catch {}


        setTimeout(
          () => {

            jobs.delete(
              jobId
            );

            cleanupDirectory(
              jobDir
            );

          },

          10 * 60 * 1000
        );


        return;

      }


      // ==================================================
      // LIBREOFFICE PROFILE
      // ==================================================

      const jobProfileDir =
        path.join(
          os.tmpdir(),
          `lo-profile-${jobId}`
        );


      fs.mkdirSync(
        jobProfileDir,
        {
          recursive:
            true
        }
      );


      let libreOfficeArgs = [

        "--headless",

        "--nologo",

        "--nodefault",

        "--nofirststartwizard",

        "--nolockcheck",

        "-env:UserInstallation=file://" +
          jobProfileDir,

        "--convert-to"

      ];


      // ==================================================
      // WORD -> PDF
      // ==================================================

      if (
        tool ===
        "word-to-pdf"
      ) {

        libreOfficeArgs.push(
          "pdf:writer_pdf_Export"
        );

      }


      // ==================================================
      // POWERPOINT -> PDF
      // ==================================================

      else if (
        tool ===
        "powerpoint-to-pdf"
      ) {

        libreOfficeArgs.push(
          "pdf:impress_pdf_Export"
        );

      }


      // ==================================================
      // EXCEL -> PDF
      // ==================================================

      else if (
        tool ===
        "excel-to-pdf"
      ) {

        libreOfficeArgs.push(
          "pdf:calc_pdf_Export"
        );

      }


      // ==================================================
      // PDF -> POWERPOINT
      // COMPRESSED VERSION
      // ==================================================

      else if (
        tool ===
        "pdf-to-powerpoint"
      ) {

        const pptTempDir =
          fs.mkdtempSync(
            path.join(
              os.tmpdir(),
              "pdf-ppt-"
            )
          );


        try {

          const prefix =
            path.join(
              pptTempDir,
              "page"
            );


          // ------------------------------------------------
          // RENDER PDF AT 120 DPI
          // ------------------------------------------------

          await new Promise(
            (
              resolve,
              reject
            ) => {

              execFile(
                "pdftoppm",

                [
                  "-png",

                  "-r",
                  "120",

                  inputPath,

                  prefix
                ],

                {
                  timeout:
                    180000,

                  maxBuffer:
                    50 * 1024 * 1024
                },

                (
                  error,
                  stdout,
                  stderr
                ) => {

                  if (
                    error
                  ) {

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


          // ------------------------------------------------
          // FIND RENDERED PAGES
          // ------------------------------------------------

          const pageFiles =
            fs
              .readdirSync(
                pptTempDir
              )
              .filter(
                file =>
                  /^page-\d+\.png$/i.test(
                    file
                  )
              )
              .sort(
                (
                  a,
                  b
                ) => {

                  const na =
                    parseInt(
                      a.match(
                        /\d+/
                      )[0],
                      10
                    );


                  const nb =
                    parseInt(
                      b.match(
                        /\d+/
                      )[0],
                      10
                    );


                  return na - nb;

                }
              );


          if (
            !pageFiles.length
          ) {

            throw new Error(
              "No PDF pages were rendered."
            );

          }


          // ------------------------------------------------
          // COMPRESS EACH PAGE TO JPEG
          // ------------------------------------------------

          const compressedFiles =
            [];


          for (
            let i = 0;
            i < pageFiles.length;
            i++
          ) {

            const pngPath =
              path.join(
                pptTempDir,
                pageFiles[i]
              );


            const jpgPath =
              path.join(
                pptTempDir,
                `compressed-${i + 1}.jpg`
              );


            await new Promise(
              (
                resolve,
                reject
              ) => {

                execFile(
                  "convert",

                  [
                    pngPath,

                    "-strip",

                    "-sampling-factor",
                    "4:2:0",

                    "-interlace",
                    "Plane",

                    "-quality",
                    "70",

                    jpgPath
                  ],

                  {
                    timeout:
                      60000,

                    maxBuffer:
                      50 * 1024 * 1024
                  },

                  (
                    error,
                    stdout,
                    stderr
                  ) => {

                    if (
                      error
                    ) {

                      reject(
                        new Error(
                          stderr?.trim() ||
                          error.message ||
                          "Image compression failed."
                        )
                      );

                      return;

                    }


                    resolve();

                  }
                );

              }
            );


            if (
              !fs.existsSync(
                jpgPath
              )
            ) {

              throw new Error(
                "Compressed image was not created."
              );

            }


            compressedFiles.push(
              jpgPath
            );

          }


          // ------------------------------------------------
          // CREATE POWERPOINT
          // ------------------------------------------------

          const pptx =
            new pptxgen();


          pptx.defineLayout({

            name:
              "PDF_PAGE",

            width:
              10,

            height:
              5.625

          });


          pptx.layout =
            "PDF_PAGE";


          pptx.author =
            "iLovePDF4";


          pptx.subject =
            "PDF to PowerPoint";


          pptx.title =
            outputFilename;


          pptx.company =
            "iLovePDF4";


          pptx.lang =
            "en-US";


          // ------------------------------------------------
          // ADD COMPRESSED IMAGES
          // ------------------------------------------------

          for (
            const jpgPath
            of compressedFiles
          ) {

            const slide =
              pptx.addSlide();


            slide.background = {

              color:
                "FFFFFF"

            };


            slide.addImage({

              path:
                jpgPath,

              x:
                0,

              y:
                0,

              w:
                10,

              h:
                5.625

            });

          }


          // ------------------------------------------------
          // WRITE PPTX
          // ------------------------------------------------

          await pptx.writeFile({

            fileName:
              outputPath

          });


          if (
            !fs.existsSync(
              outputPath
            )
          ) {

            throw new Error(
              "PowerPoint output was not created."
            );

          }


          const stat =
            fs.statSync(
              outputPath
            );


          if (
            !stat.isFile() ||
            stat.size < 10000
          ) {

            throw new Error(
              "PowerPoint output is invalid."
            );

          }


          jobs.set(
            jobId,
            {

              status:
                "finished",

              outputPath:
                outputPath,

              filename:
                outputFilename,

              error:
                null

            }
          );


          console.log(
            "PDF -> POWERPOINT FINISHED:",
            outputPath
          );


          console.log(
            "PPTX size:",
            stat.size,
            "bytes"
          );


        } catch (
          error
        ) {

          console.error(
            "PDF -> POWERPOINT error:",
            error
          );


          jobs.set(
            jobId,
            {

              status:
                "error",

              outputPath:
                null,

              filename:
                outputFilename,

              error:
                error.message ||
                "PDF to PowerPoint conversion failed."

            }
          );


        } finally {

          cleanupDirectory(
            pptTempDir
          );

        }


        try {

          fs.unlinkSync(
            inputPath
          );

        } catch {}


        setTimeout(
          () => {

            jobs.delete(
              jobId
            );

            cleanupDirectory(
              jobDir
            );

          },

          10 * 60 * 1000
        );


        return;

      }


      // ==================================================
      // PDF -> EXCEL
      // ==================================================

      else if (
        tool ===
        "pdf-to-excel"
      ) {

        libreOfficeArgs.push(
          "xlsx:Calc MS Excel 2007 XML"
        );


        libreOfficeArgs.push(
          "--infilter=draw_pdf_import"
        );

      }


      // ==================================================
      // UNKNOWN TOOL
      // ==================================================

      else {

        jobs.set(
          jobId,
          {

            status:
              "error",

            outputPath:
              null,

            filename:
              outputFilename,

            error:
              "Unsupported conversion tool."

          }
        );


        try {

          fs.unlinkSync(
            inputPath
          );

        } catch {}


        return;

      }
