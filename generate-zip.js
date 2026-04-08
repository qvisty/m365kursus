const fs = require("fs");
const path = require("path");
const JSZip = require("jszip");

const DOCX_DIR = path.join(__dirname, "vejledninger", "docx");
const ZIP_PATH = path.join(DOCX_DIR, "m365-alle-vejledninger.zip");

async function main() {
  const zip = new JSZip();
  const files = fs.readdirSync(DOCX_DIR).filter(f => f.endsWith(".docx")).sort();

  for (const file of files) {
    const data = fs.readFileSync(path.join(DOCX_DIR, file));
    zip.file(file, data);
  }

  const buffer = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  fs.writeFileSync(ZIP_PATH, buffer);
  console.log(`Created ${ZIP_PATH} with ${files.length} files`);
}

main().catch(err => { console.error(err); process.exit(1); });
