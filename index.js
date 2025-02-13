const fs = require("fs").promises;
const fsp = require('fs')
const path = require("path");
const os = require("os");
const PdfParse = require("pdf-parse");
const ps = require('prompt-sync');
const textract = require('textract')
const mammoth = require('mammoth');
const unzipper = require("unzipper");
const xml2js = require('xml2js')


const MAX_FILES = 10000;
const EXCLUDED_DIRS = ["node_modules", ".git", "vendor", "build"]; // Add more if needed
let fileCount = 0; // Track the number of printed files

let entryPath = path.join(os.homedir(), "OneDrive", "Desktop");
let destinationFolderPath = null;

let keyWords = ["groupstudy"]


//Function to create new folder
const createFolder = async (newFolderPath) => {
    try {
        await fs.mkdir(newFolderPath); // Creates folder, recursive allows creating nested folders
        console.log(`Folder created at: ${newFolderPath}`);
    } catch (error) {
        console.error("Error creating folder:", error.message);
    }
};




//Function to check and delete folder if it is empty
const checkIfFolderIsEmptyIfSoDeleteIt = async(folderPath)=>{
    try {
        await fs.rmdir(folderPath);
        console.log(folderPath, " deletede successfully\n");
        
    } catch (error) {
        // console.log(`The ${folderPath} is not empty so i can't delete it.`);        
    }
}


//function to move file from source to destination
const moveFilesFromSourceToDestination = async(sourcePath, destinationPath)=>{

    destinationPath = path.join(destinationPath, path.basename(sourcePath));

    try {
        await fs.chmod(sourcePath, 0o666); // Remove read-only restrictions
        await fs.rename(sourcePath, destinationPath);
        console.log(`${sourcePath} ----> ${destinationFolderPath}\n`);
        
    } catch (error) {
        console.log(error.message);
    }
    
}



const handleCurrentPdfFile = async(filePath)=>{
    
    try {
        const dataBuffer = await fs.readFile(filePath);
        const data = (await PdfParse(dataBuffer)).text.toLowerCase();
        
        for(let key of keyWords){
            if(data.includes(key.toLowerCase())){
                await moveFilesFromSourceToDestination(filePath, destinationFolderPath);
            }
        }
        
    } catch (error) {
        console.log(error.message);
    }
}



const handleCurrentDocxFile = async(filePath)=>{
    try {
        const buffer = await fs.readFile(filePath);
        const { value: text } = await mammoth.extractRawText({ buffer });

        const data = text;
 
        for(let key of keyWords){
            if(data.includes(key.toLowerCase())){
                await moveFilesFromSourceToDestination(filePath, destinationFolderPath);
            }
        }

    } catch (error) {
        console.log("f***\n", error.message);
    }
}


async function extractTextFromPPTX(filePath) {
    const pptxText = [];

    // Unzip PPTX file
    const zip = fsp.createReadStream(filePath).pipe(unzipper.Parse({ forceStream: true }));

    for await (const entry of zip) {
        // Look for slide XML files
        if (entry.path.startsWith("ppt/slides/slide") && entry.path.endsWith(".xml")) {
            const content = await entry.buffer();
            const parsedXml = await xml2js.parseStringPromise(content);
            
            // Extract text from slide
            const textElements = extractText(parsedXml);
            pptxText.push(...textElements);
        } else {
            entry.autodrain();
        }
    }

    // console.log("Extracted Text:", pptxText.join(" "));
    return pptxText.join(" ");
}

// Function to extract text from XML structure
function extractText(xmlObj) {
    let textArray = [];

    function traverse(obj) {
        if (typeof obj === "object") {
            for (const key in obj) {
                if (key === "a:t") {
                    textArray.push(obj[key][0]); // Extract actual text
                } else {
                    traverse(obj[key]);
                }
            }
        }
    }

    traverse(xmlObj);
    return textArray;
}

// Run function

const handleCurrentPptxFile = async(filePath)=>{
    const data = (await extractTextFromPPTX(filePath)).toLowerCase();
    
    for(let key of keyWords){
        if(data.includes(key.toLowerCase())){
            await moveFilesFromSourceToDestination(filePath, destinationFolderPath);
        }
    }
    
}



const readFilesRecursively = async (dir) => {

    try {
        // Stop recursion if we have printed enough files
        if (fileCount >= MAX_FILES) return;

        const entries = await fs.readdir(dir, { withFileTypes: true });
   
        for (const entry of entries) {
            const fullPath = path.join(dir, entry.name);

            // Skip large or unnecessary directories
            if (entry.isDirectory() && EXCLUDED_DIRS.includes(entry.name)) {
                // console.log(`Skipping folder: ${fullPath}`);
                return;
            }

            if (entry.isDirectory()) {
                await readFilesRecursively(fullPath);
            } else if (entry.isFile()) {

                if((entry.name.endsWith(".pdf"))){
                    await handleCurrentPdfFile(fullPath);
                } else if(entry.name.endsWith(".docx")){
                    await handleCurrentDocxFile(fullPath);
                } else if(entry.name.endsWith(".pptx")){
                    await handleCurrentPptxFile(fullPath);
                }

                fileCount++;
                if (fileCount >= MAX_FILES) return; // Stop if we reach 500
            }
        }
        // await checkIfFolderIsEmptyIfSoDeleteIt(dir);
    } catch (error) {
        console.error(`Error reading ${dir}:`, error.message);
    }
};


const readFiles = async (entryPath) => {
    try {
        // Check if OneDrive directory exists
        const stat = await fs.stat(entryPath);
        if (!stat.isDirectory()) {
            console.log("OneDrive directory not found.");
            return;
        }

        console.log(`Scanning for PDFs in: ${entryPath}`);
        await readFilesRecursively(entryPath);

        console.log(`Printed ${fileCount} PDF file paths.`);
    } catch (error) {
        console.error("Error:", error.message);
    }
};


const resolvePath = async(folderPath)=>{
    folderPath = folderPath.replace(/^"(.*)"$/, '$1');
    folderPath = path.resolve(folderPath);
    
    return folderPath;
}



const main = async()=>{
    const prompt = ps();
    keyWords =  prompt("Enter key words: ")
    keyWords = keyWords.split(',');
    
    destinationFolderPath = prompt("Enter destination path: ");
    destinationFolderPath = await resolvePath(destinationFolderPath);

    const newFolderName = prompt("Enter new folder name: ");
    destinationFolderPath = path.join(destinationFolderPath, newFolderName);
    
    entryPath = prompt("Enter folder path where you have to combine similar files: ");
    entryPath = await resolvePath(entryPath);
    // console.log(destinationFolderPath);
    
    await createFolder(destinationFolderPath);
    await readFiles(entryPath);

}

main()
//C:\Users\Sachin\OneDrive\Desktop\ESA\iotplusml\ss