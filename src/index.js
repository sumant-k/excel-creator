const ExcelJS = require("exceljs");
const fs = require("fs");
const image2base64 = require("image-to-base64"); // You can use this library to convert images to base64
const path = require("path");

// Create a new Excel workbook and worksheet
const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet("Images");

// Set the column headers
worksheet.columns = [
  { header: "Tag", key: "tag" },
  { header: "Image", key: "image" },
];

// get File extension from file path
const getFileExtension = (filePath) => {
  try {
    return filePath.split(".").pop();
  } catch (error) {
    return "png";
  }
};

const createDirIfNotExists = (dir) => (!fs.existsSync(dir) ? fs.mkdirSync(dir) : undefined);

// Generate the Excel file
const writeDataToFile = async (imageList) => {
  try {
    for (let i = 0; i < imageList.length; i++) {
      const imageEntry = imageList[i];
      const fileExtension = getFileExtension(imageEntry.path);
      // Convert the image to base64
      // Add a row to the worksheet with the image name and base64 image data
      const imageId = workbook.addImage({
        filename: imageEntry.path,
        extension: fileExtension,
      });
      // const opt = {
      //   tl: { col: 0, row: 0 },
      //   ext: { width: 100, height: 50 },
      // };
      const addingImage = worksheet.addImage(imageId, `B${i + 2}:B${i + 3}`);
      console.log("imageId1 ", addingImage);
      worksheet.addRow({
        tag: imageEntry.tag,
        image: "",
      });
    }
    //
    const outputDirectory = path.join(__dirname, "../output");
    createDirIfNotExists(outputDirectory);
    const response = workbook.xlsx.writeFile(outputDirectory + "/image_excel.xlsx");
    console.log("Successfully written to file", response);
  } catch (error) {
    console.log("Error writing to file", error);
  }
};

const getImageList = () => {
  // get all files in the ./resource directory
  const resourceDirectory = path.join(__dirname, "../resource");
  const files = fs.readdirSync(resourceDirectory);
  return files.map((file) => {
    return { tag: file, path: path.join(resourceDirectory, file) };
  });
};

const files = getImageList();
console.log("getImageList ", writeDataToFile(files));
