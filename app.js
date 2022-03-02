const readXlsxFile = require("read-excel-file/node");

const express = require("express");
const app = express();
const port = 3000;

const readExcel = async () => {
  const rows = await readXlsxFile("./file2.xlsx");
  return rows
    .map((row) => row.filter((cell) => cell !== null))
    .filter((row) => row.length > 0);
};

const extractValues = async () => {
  const allRows = await readExcel();

  const dataMapped = allRows.map((row, index) => {
    const hasContactNum = row.indexOf("Contract No") > -1;
    if (hasContactNum) {
      //console.log(row);
      const contractNo = allRows[index + 1][2];
      const sizesT1 = allRows[index].slice(7);
      const quantitiesT1 = allRows[index + 1].slice(7);
      const sizesT2 = allRows[index + 3].includes("Size")
        ? allRows[index + 3].slice(1)
        : [];
      const quantitiesT2 = allRows[index + 4].includes("Quantity")
        ? allRows[index + 4].slice(1)
        : [];

      const sizesAndQuantitiesT1 = sizesT1.map((size, index) => [
        size,
        quantitiesT1[index],
      ]);

      const sizesAndQuantitiesT2 = sizesT2.map((size, index) => [
        size,
        quantitiesT2[index],
      ]);

      const contractObj = {
        contractNo: allRows[index + 1][2],
        sizesT1,
        quantitiesT1,
        sizesT2,
        quantitiesT2,
        sizesAndQuantities: [...sizesAndQuantitiesT1, ...sizesAndQuantitiesT2],
      };

      return contractObj;
    }

    return null;
  });

  return dataMapped.filter((row) => row !== null);
};

app.get("/", async (req, res) => {
  const data = await extractValues();
  res.json(data);
});

app.listen(port, () => {
  console.log(`Example app listening on port ${port}`);
});
