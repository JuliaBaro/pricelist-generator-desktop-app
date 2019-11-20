var Excel = require("exceljs");

//Global variables:
let mainColumnArr;
let mainRowArr;
let readWorksheet;
let idsArrFSB_SB = [];
var errorList=[];

let codeStart1 = "FSB.20.";
let codeStart2 = "SB.20.";

let projectProductIdArr;
let quantityArr;
var projectWorksheet;
let idsAndQuantityArr = [];

let arrOfAllElements = [];

let sum = [];

let unitPriceArr = [];
let fullPriceArr = [];
let arrOfProjectColumns = [];
var flawHappened=0;

function Init()
{
    flawHappened=0;
    idsArrFSB_SB = [];
    errorList=[];
    idsAndQuantityArr = [];    
    arrOfAllElements = [];
    sum = [];    
    unitPriceArr = [];
    fullPriceArr = [];
    arrOfProjectColumns = [];
}

//Checks id-s
function checkId(productid)
{
    if (productid.match(/^[F]?SB.([0-9]{2}).([0-9]{4}).([0-9]{4}).([0-9]{2})$/))
    {
        return true;
    }
    else
    {
        return false;
    }
}

//Checks if a cells type is string or number.
function checkCellType(cell)
{
    if (typeof cell === "string" || typeof cell === "number")
    {
        return true;
    }
    else
    {
        console.log("false " + typeof cell + " " + cell);
        return false;
    }
}

//This function copies every Project.xlsx column as an array into an array like so [[][][]]
function readProject(project)
{
    if(flawHappened)
    {
        return Promise.reject("flawHappened");
    }
    let readWorkbook = new Excel.Workbook();
    return readWorkbook.xlsx.readFile(project)
    .then(function()
    {   
        let columnIndexArr = ["A", "B", "C", "D", "E", "F", "G", "H"]; 

        for (let j = 0; j < columnIndexArr.length; j++)
        {
            let worksheet = readWorkbook.getWorksheet('Matten');
            let nameColArr = worksheet.getColumn(columnIndexArr[j]).values;
            for(let k=0;k<nameColArr.length;k++)
            {
                if(k>3 && !checkCellType(nameColArr[k]))
                {
                    errorList.push("type error in "+project+" col "+j+", row "+k);
                }
            }
            arrOfProjectColumns.push(nameColArr);
        }

        if(errorList.length>0)
        {
            //Here throw exception
            /* for(let i=0;i<errorList.length;i++)
            {
                console.log(errorList[i]);
            }*/
            flawHappened=1;
            console.log("readProject failed");
            throw "readProject failed in "+project;
        }
        //console.log(arrOfProjectColumns);
    }).catch((error)=>handleError(error))
}

//Writes xlsx - copies Project.xlsx + unit price column + quantity based price column + final price.
function mergedResult(project, result)
{
    console.log("reading project file");
    if(flawHappened)
    {
        return Promise.reject("flawHappened");
    }

    return readProject(project)
    .then(function()
    {
        if(flawHappened)
        {
            console.log("flawHappened, skip result writing");
            return;
        }
        let mergeWorkbook = new Excel.Workbook();
        let mergeWorksheet = mergeWorkbook.addWorksheet('Matten');

        let columnIndexArr1 = ["A", "B", "C", "D", "E", "F", "G", "H"]; 
        let columnIndexArr2 = ["I", "J", "K"];

        for (let i = 0; i < columnIndexArr1.length; i++)
        {
            mergeWorksheet.getColumn(columnIndexArr1[i]).values = arrOfProjectColumns[i];
        }

        //console.log("unitPriceArr.length elotte: " + unitPriceArr.length);
        //console.log("fullPriceArr.length elotte: " + fullPriceArr.length);

        for (let i = 0; i < arrOfAllElements.length; i++)
        {
            unitPriceArr.push(arrOfAllElements[i].unitPrice);
            fullPriceArr.push(arrOfAllElements[i].fullPrice);
        }

        //console.log("unitPriceArr.length utana: " + unitPriceArr.length);
        //console.log("fullPriceArr.length utana: " + fullPriceArr.length);

        unitPriceArr.unshift("", "", "", "Einheitspreise");
        fullPriceArr.unshift("", "", "", "Gesamtpreis"); 

        //console.log("unitPriceArr.length utana2: " + unitPriceArr.length);
        //console.log("fullPriceArr.length utana2: " + fullPriceArr.length);
        
        sum.unshift("", "", "", "Totalpreis");

        mergeWorksheet.getColumn(columnIndexArr2[0]).values = unitPriceArr;
        //console.log("arrOfAllElements.length mergedResult: " + arrOfAllElements.length);
        mergeWorksheet.getColumn(columnIndexArr2[1]).values = fullPriceArr;
        mergeWorksheet.getColumn(columnIndexArr2[2]).values = sum;

        return mergeWorkbook.xlsx.writeFile(result)
        .then(function()
        {
            console.log("result file written "+flawHappened);
            /*console.log(mergeWorksheet.getColumn(columnIndexArr1[1]).values);
            console.log(mergeWorksheet.getColumn(columnIndexArr2[0]).values);
            console.log(mergeWorksheet.getColumn(columnIndexArr2[1]).values);
            console.log(mergeWorksheet.getColumn(columnIndexArr2[2]).values);*/
        })
    })
}

//This Function creates id-s based on FSB/SB.xlsx files and assigns corresponding unit prices.
function pricelistReader(FSB_SB, codeStart) 
{
    if(flawHappened)
    {
        return Promise.reject("flawHappened");
    }

    let workbook = new Excel.Workbook();
    let countFSB_SB = 0;
    return ((workbook.xlsx.readFile(FSB_SB)
    .then(function()
    {
        console.log("pricelistReader: " + FSB_SB);
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
                if(!checkCellType(mainRowArr[j]))
                {
                    errorList.push("type error in "+FSB_SB+" at row "+j);
                }
                if(!checkCellType(mainColumnArr[i]))
                {
                    errorList.push("type error in "+FSB_SB+" at column "+i);
                }
                if(!checkCellType(oneLine[j]))
                {
                    errorList.push("type error in "+FSB_SB+" at cell value "+oneLine[j]);
                }
                countFSB_SB = countFSB_SB + 1;
                let euro = oneLine[j];
                let productid = (codeStart + mainRowArr[j] + "." + mainColumnArr[i] + ".00");
                if(!checkId(productid))
                {
                    errorList.push("error in "+FSB_SB+" at productid "+productid);
                }
                idsArrFSB_SB.push({productid, euro});
                //console.log({productid, euro});
                //console.log(countFSB_SB);
            }
        }
        if(errorList.length>0)
        {
            //Here throw exception
            //return Promise.reject("pricelistReader failed in "+FSB_SB);
            flawHappened=1;
            throw "pricelistReader failed in "+FSB_SB;
        }
    }).catch(error=>handleError(error))))
}

//This function calls the above function with two different parameters (FSB + SB).
function readFSB_SB(FSB, SB)
{
    console.log('reading FSB and SB');
    return pricelistReader(FSB, codeStart1)
    .then(pricelistReader(SB, codeStart2))
    .catch((error)=>handleError(error));
}

function project(FSB, SB, project)
{
    console.log('reading Project.xlsx');
    Init();
    let workbook = new Excel.Workbook();

    return readFSB_SB(FSB, SB)
    .then(workbook.xlsx.readFile(project)
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
                
                if(!checkCellType(value))
                {
                    errorList.push("type error in "+project+" at quantity row "+i);
                }
                if(!checkId(id))
                {
                    errorList.push("error in " +project+" at "+id);
                }
                //Stores id-s and quantities in an object.
            }

            if(errorList.length>0)
            {
                //Here throw exception
                flawHappened=1;
                console.log("project failed");
                throw "project reader failed in "+project;
            }            
            //console.log('idsAndQuantityArr.length: '+idsAndQuantityArr.length);

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
            //console.log("arrOfAllElements.length array KESZ: " + arrOfAllElements.length);
            //console.log("osszes adat: " + arrOfAllElements);

            //Writes out the sum of fullPrices. This is the final price that the client should pay.
            let sumFullPrice = 0;

            for (let i = 0; i < arrOfAllElements.length; i++)
            {
                sumFullPrice = sumFullPrice + arrOfAllElements[i].fullPrice;
            }
            sum.push(sumFullPrice.toFixed(3));
            console.log("Total Preis: " + sum + " €");
        }).catch((error)=>handleError(error))
    )
}

/*
project("FSB.xlsx", "SB.xlsx", "Project.xlsx")
.then(mergedResult("Project.xlsx", "Result.xlsx"))
.catch(function(error) {
    console.log("catch: "+error);
    for(let i=0;i<errorList.length;i++)
    {
        console.log(errorList[i]);
    }
  });*/

function handleError(error)
{
    console.log("!!!!!!!!! catch: "+error);
    for(let i=0;i<errorList.length;i++)
    {
        console.log(errorList[i]);
    }
    flawHappened=1;
}

project("FSB.xlsx", "SB.xlsx", "Project_test_failed.xlsx")
.then(mergedResult("Project_test_failed.xlsx", "Result.xlsx"))
.catch((error)=>handleError(error));
  