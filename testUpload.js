require("dotenv").config();

const { uploadToOneDrive } = require("./onedrive.service");

async function test() {
  const company = process.env.ONEDRIVE_TEST_COMPANY || "DFP";
  const baseFolder = process.env.ONEDRIVE_SALES_FOLDER || "Documents/test/ventas";

  const result = await uploadToOneDrive(
    "./test.xlsx",
    `VENTAS PRUEBA ${company}.xlsx`,
    `${baseFolder}/${company}`
  );

  console.log("Subido a OneDrive:");
  console.log(result.webUrl);
}

test().catch(console.error);
