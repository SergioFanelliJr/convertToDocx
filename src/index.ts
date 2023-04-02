const fs = require("node:fs/promises");
const fsSync = require("node:fs");
const childProcess = require("node:child_process");
const path = require("node:path");
const os = require("node:os");

const extIn = ".doc";
const extOut = "docx";
const cwd = process.cwd();
const _os = os.platform();
const sofficePath = "C:\\Program Files\\LibreOffice\\program\\soffice.exe";

async function preFlight() {
    console.clear();
    console.info("\x1b[34m" + "Elefante.dev" + "\x1b[0m");
    console.info("\x1b[32m" + "OS: " + "\x1b[0m" + `${_os}`);
    console.info("\x1b[32m" + "Current directory: " + "\x1b[0m" + `${cwd}`);
    if (_os !== "win32") {
        console.error("\x1b[31m" + "This script is only for Windows" + "\x1b[0m");
        return false;
    }
    const filesToConvert = await fs.readdir(cwd);
    if (filesToConvert.length === 0) {
        console.error("\x1b[31m" + "No files to convert" + "\x1b[0m");
        return false;
    }
    const filesToConvertFiltered = filesToConvert.filter((file: any) => {
        return file.endsWith(extIn);
    });
    console.info("----------------------------------------");
    console.info(
        "\x1b[32m" + `There is ${filesToConvertFiltered.length} files to convert` + "\x1b[0m"
    );
    console.info("----------------------------------------");

    if (fsSync.existsSync(sofficePath) === false) {
        console.error("\x1b[31m" + "LibreOffice not found" + "\x1b[0m");
        return false;
    }
    console.info("LibreOffice found");
    console.info("----------------------------------------");

    if (!fsSync.existsSync(path.join(cwd, "converted"))) {
        await fs.mkdir(path.join(cwd, "converted"));
    }
    return filesToConvertFiltered;
}

async function convertFiles() {
    const filesToConvert = await preFlight();
    if (filesToConvert === false) {
        console.error("\x1b[31m" + "Pre-flight failed" + "\x1b[0m");
        return;
    }
    for (const file of filesToConvert) {
        const fileIn = path.join(cwd, file);
        const command = `"${sofficePath}" --headless --convert-to ${extOut} --outdir "${path.join(
            cwd,
            "converted"
        )}" "${fileIn}"`;
        console.info("\x1b[33m" + "Converting -> |" + "\x1b[0m" + `${fileIn}` + "\x1b[33m");
        childProcess.execSync(command);
        console.log("\x1b[32m" + "Converted <-  |" + "\x1b[0m" + `${fileIn}`);
    }
    console.info("----------------------------------------");
}

convertFiles();
