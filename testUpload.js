const { uploadToOneDrive } = require("./onedrive.service");

async function test() {
  const result = await uploadToOneDrive(
    "./test.xlsx",
    "test.xlsx",
    "Documents/test"
  );

  console.log("🔥 Subido:");
  console.log(result.webUrl);
}

test().catch(console.error);