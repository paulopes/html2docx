#!/usr/bin/env node

process.noDeprecation = true;

const { program } = require("commander");
const fs = require("fs");
const path = require("path");
const HTMLtoDOCX = require("html-to-docx");

function readFileOrInline(value) {
  const resolved = path.resolve(value);
  if (fs.existsSync(resolved)) {
    return fs.readFileSync(resolved, "utf-8");
  }
  return value;
}

program
  .name("html2docx")
  .description("Convert an HTML file to a DOCX file")
  .version("1.0.0")
  .argument("<input>", "path to the HTML file")
  .option("-o, --output <path>", "output DOCX file path")
  .option("--landscape", "landscape orientation (default: portrait)")
  .option("--margin <inches>", "uniform page margin in inches", parseFloat)
  .option("--margin-top <inches>", "top margin in inches", parseFloat)
  .option("--margin-bottom <inches>", "bottom margin in inches", parseFloat)
  .option("--margin-left <inches>", "left margin in inches", parseFloat)
  .option("--margin-right <inches>", "right margin in inches", parseFloat)
  .option("--font <name>", "default font name")
  .option("--font-size <points>", "default font size in points", parseFloat)
  .option("--title <text>", "document title metadata")
  .option("--creator <text>", "document creator metadata")
  .option("--header <file-or-html>", "header HTML (file path or inline HTML string)")
  .option("--footer-html <file-or-html>", "footer HTML (file path or inline HTML string)")
  .option("--no-page-numbers", "disable page numbers in footer")
  .option("--line-numbers", "enable line numbering")
  .option("--lang <code>", "language code for spell check (default: en-US)")
  .action(async (input, options) => {
    const inputPath = path.resolve(input);

    if (!fs.existsSync(inputPath)) {
      console.error(`Error: file not found: ${inputPath}`);
      process.exit(1);
    }

    const outputPath = options.output
      ? path.resolve(options.output)
      : inputPath.replace(/\.html?$/i, "") + ".docx";

    const html = fs.readFileSync(inputPath, "utf-8");

    // Build margins object — specific edges override the uniform --margin value
    const inToTwip = (inches) => `${inches}in`;
    const margins = {};
    if (options.margin != null) {
      const val = inToTwip(options.margin);
      Object.assign(margins, { top: val, right: val, bottom: val, left: val });
    }
    if (options.marginTop != null) margins.top = inToTwip(options.marginTop);
    if (options.marginBottom != null) margins.bottom = inToTwip(options.marginBottom);
    if (options.marginLeft != null) margins.left = inToTwip(options.marginLeft);
    if (options.marginRight != null) margins.right = inToTwip(options.marginRight);

    // Build document options
    const docOptions = {
      table: { row: { cantSplit: true } },
      footer: true,
      pageNumber: options.pageNumbers !== false,
    };

    if (options.landscape) docOptions.orientation = "landscape";
    if (Object.keys(margins).length) docOptions.margins = margins;
    if (options.font) docOptions.font = options.font;
    if (options.fontSize) docOptions.fontSize = `${options.fontSize}pt`;
    if (options.title) docOptions.title = options.title;
    if (options.creator) docOptions.creator = options.creator;
    if (options.lineNumbers) docOptions.lineNumber = true;
    if (options.lang) docOptions.lang = options.lang;
    if (options.header) docOptions.header = true;

    // Header/footer HTML strings (2nd and 4th arguments to HTMLtoDOCX)
    const headerHTML = options.header ? readFileOrInline(options.header) : null;
    const footerHTML = options.footerHtml ? readFileOrInline(options.footerHtml) : undefined;

    try {
      const buffer = await HTMLtoDOCX(html, headerHTML, docOptions, footerHTML);

      fs.writeFileSync(outputPath, buffer);
      console.log(`Created: ${outputPath}`);
    } catch (err) {
      console.error(`Conversion failed: ${err.message}`);
      process.exit(1);
    }
  });

program.parse();
