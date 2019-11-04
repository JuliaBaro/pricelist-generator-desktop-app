var Excel = require("exceljs");
var workbook = new Excel.Workbook();

let mainColumnArr;
let mainRowArr;
let readWorksheet;
let idsArrFSB_SB = [];

let codeStart1 = "FSB.20.";
let codeStart2 = "SB.20.";

let projectProductIdArr;
let quantityArr;
var projectWorksheet;
let idsAndQuantityArr = [];

let arrOfAllElements = [];
let fullPrice;

function checkId(productid)
{
    if (productid.match(/^[F]?SB.([0-9]{2}).([0-9]{4}).([0-9]{4}).([0-9]{2})$/))
    {
        console.log("true");
        return true;
    }
    else
    {
        console.log("Erno .catch");
        return false;
    }
}
//checkId("FSB.20.3000.0600.00");
//checkId("SB.20.1500.0150.00");
//checkId("S.20.1500.0150.00");

/*function checkCellType()
{
    for(let i = 0; i < .length; i++)
    {
        for (let j = 0; j < .length; j++)
        {
            if(typeof cells === "string")
            {
                console.log("true");
                return true;
            }
            else 
            {
                console.log("Erno .catch");
                return false;
            }
        }
    } 
}
checkCellType();*/

function pricelistReader(FSB, codeStart1) 
{
    workbook.xlsx.readFile(FSB)
    .then(function()
        {
        //Reading id-s from FSB and SB
        readWorksheet = workbook.getWorksheet('Munka1');
        mainColumnArr = readWorksheet.getColumn('A').values;    
        mainRowArr = readWorksheet.getRow(2).values;

        for (let i = 2; i < mainColumnArr.length; i++) 
        {
            if (mainColumnArr[i] < 1000) 
            {
                mainColumnArr[i] = "0" + mainColumnArr[i];
            }
        }
        readWorksheet = workbook.getWorksheet('Munka1');
        for (let i = 3; i < mainColumnArr.length; i++) 
        {
            let oneLine = readWorksheet.getRow(i).values;
            for (let j = 2; j < mainRowArr.length-1; j++) 
            {
                let euro = oneLine[j];
                let productid = (codeStart1 + mainRowArr[j] + "." + mainColumnArr[i] + ".00");
                //var productid = ("FSB.20." + mainRowArr[j] + "." + mainColumnArr[i] + ".00");
                    checkId(productid);
                idsArrFSB_SB.push({productid, euro});
            }
        }
        console.log(idsArrFSB_SB);
        return idsArrFSB_SB;
    })
}

function readFSB_SB (FSB, SB)
{
    //console.log('reading FSB and SB');
    workbook.xlsx.readFile(FSB)
    .then (function()
    {
        return pricelistReader(FSB, codeStart1); 
    })
    .then (function()
    {
        return pricelistReader(SB, codeStart2);
    })
}

function project(FSB, SB, project)
{
    //console.log('reading Project.xlsx');
    readFSB_SB(FSB, SB);//then

    workbook.xlsx.readFile(project)
    .then(function() 
    {
        //Reads id and quantity coulumn from Project.xlsx.
        projectWorksheet = workbook.getWorksheet('Matten');
        projectProductIdArr = projectWorksheet.getColumn('A').values;  
        quantityArr = projectWorksheet.getColumn('C').values;

        for (let i = 5; i < projectProductIdArr.length; i++) 
        {
            var id = projectProductIdArr[i];
            var value = quantityArr[i];
            idsAndQuantityArr.push({id, value});
                checkId(id);
            //Stores id-s and quantities in an object.
        }
        console.log(idsAndQuantityArr);
        //return idsAndQuantityArr;

        //Looks for Project.xlsx id-s in the array based on FSB/SB.xlsx.
        console.log('final array with all data');
        for (let i = 0; i < idsAndQuantityArr.length; i++)
        {
            for (let j = 0; j < idsArrFSB_SB.length; j++)
            {
                if (idsAndQuantityArr[i].id === idsArrFSB_SB[j].productid)
                {
                    let id = idsAndQuantityArr[i].id;
                    let quantity = Number(idsAndQuantityArr[i].value);
                    let unitPrice = Number(idsArrFSB_SB[j].euro);
                    let quantityTimesunitPrice = quantity * unitPrice;
                    fullPrice = Number(quantityTimesunitPrice.toFixed(3));
                    arrOfAllElements.push({id, quantity, unitPrice, fullPrice});
                }
            }
        }
        console.log(arrOfAllElements);

        //Writes out the sum of fullPrices. This is the final price that the client should pay.
        console.log('Total price');
        let sumFullPrice = 0;

        for (let i = 0; i < arrOfAllElements.length; i++)
        {
            sumFullPrice = sumFullPrice + arrOfAllElements[i].fullPrice;
        }
        let sum = sumFullPrice.toFixed(3);
        console.log("Total Preis: " + sum + " €");
        //return
    })
}

/*function writeResult()
{
    let columnIndex = ["A", "B", "C", "D", "E", "F", "G", "H"];
    
    for (let i = 0; i < columnsIndex.length; i++)
    {
        let column = getColumn(columnIndex[i]).values;
        console.log(column);
    }
}*/

//readFSB_SB("FSB.xlsx", "SB.xlsx");
project("FSB.xlsx", "SB.xlsx", "Project.xlsx");
//writeResult("Result.xlsx");

//------------------------------------------------------------------------------------------

//hint for main function with promises

//1: read files > FSB, SB, Project
//2: write file > Result

/*
readFSB_SB(FSB, SB)
{
    readWorkbook.readfile(FSB)
    .then(pricelistReader(FSB, codeStart))
    .then(pricelistReader(SB, codeStart));
    //return
}

readProject(project, FSB, SB)
{
    readFSB_SB(FSB, SB)
    .then //readProject internal code below
    //return
}

readProject("FSB.xlsx", "SB.xlsx", "Project.xlsx")
.then //writeResult("Result.xlsx")

//writeResult hint
let columnIndex = ["A", "B", "C", "D", "E", "F", "G", "H"];
getColumn(columnIndex[i]).values;
*/

//------------------------------------------------------------------------------------------

//id checker
/*
function checkId(productid)
{
    if (productid.match(/^[F]?SB.([0-9]{2}).([0-9]{4}).([0-9]{4}).([0-9]{2})$/))
    {
        console.log("true");
        return true;
    }
    else
    {
        console.log("Erno .catch");
        return false;
    }
}
checkId("FSB.20.3000.0600.00");
checkId("SB.20.1500.0150.00");
checkId("S.20.1500.0150.00");
*/

//------------------------------------------------------------------------------------------

//cell type checker (should be string)
/*function checkCellType()
{


    if(typeof cells === "string")
    {
        console.log("true");
        return true;
    }
    else 
    {
        console.log("Erno .catch");
        return false;
    }
}
checkCellType();
*/


