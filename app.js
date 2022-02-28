const readXlsxFile = require("read-excel-file/node");

const readExcel = async () => {
  const rows = await readXlsxFile("./file2.xlsx");
  return rows
    .map((row) => row.filter((cell) => cell !== null))
    .filter((row) => row.length > 0);
};

const extractValues = async () => {
  const allRows = await readExcel();

  allRows.filter((intArr, index) => {
    const hasContractNo = intArr.indexOf("Contract No") > -1;

    if (hasContractNo) {
      //console.log(allRows[index + 1]);

      const sizesT1 = allRows[index].slice(7);
      const quantities = allRows[index + 1].slice(7);
      //console.log("Sizes2", sizesT2);
      //console.log("Quantities", quantities);
      const sizesAndQuantities = sizesT1.map((size, index) => [
        size,
        quantities[index],
      ]);
      const contractObj = {
        contractNo: allRows[index + 1][2],
        sizesAndQuantities,
      };
      //console.log("SizesAndQuantities", sizesAndQuantities);
      console.log(contractObj);
    }

    return hasContractNo;
  });

  //console.log(allRows);
};

extractValues();
