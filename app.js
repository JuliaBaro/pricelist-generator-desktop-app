//Global variables.
var Excel = require("exceljs");
var workbook = new Excel.Workbook();

var mainColumn;
var mainRow;
var worksheet;
var itemNoArray = [];

let projectProductId;
let quantity;
let idsAndQuantityArray = [];
var projectWorksheet;

let arrOfAllElements = [];

//Checks id.
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

/*checkId("FSB.20.3000.0600.00");
checkId("SB.20.1500.0150.00");
checkId("S.20.1500.0150.00");*/

//Read SB and/or FSB.
console.log('Reading FSB.xlsx');
function readMatrix(inputFileName) //fsbFileName, sbFileName - 2 input
{
    let inputFileName; //SB and/or FSB
    let codeStart; //"FSB.20." and/or "SB.20."

    if (inputFileName == 'FSB.xlsx')
    {
        codeStart == "FSB.20.";
    }
    else if (inputFileName == 'SB.xlsx')
    {
        codeStart == "SB.20.";
    }

    //workbook.xlsx.readFile(inputFileName)
    workbook.xlsx.readFile('FSB.xlsx')                                         
    .then(function() 
    {
        //Reads from FSB/SB.xlsx id-s and unit prices.
        worksheet = workbook.getWorksheet('Munka1');
        mainColumn = worksheet.getColumn('A').values;    
        mainRow = worksheet.getRow(2).values;

        for (let i = 2; i < mainColumn.length; i++) 
        {
            if (mainColumn[i] < 1000) 
            {
                mainColumn[i] = "0" + mainColumn[i];
            }
        }

        worksheet = workbook.getWorksheet('Munka1');
        for (let i = 3; i < mainColumn.length; i++) 
        {
            var egysor=worksheet.getRow(i).values;
            for (let j = 2; j < mainRow.length-1; j++) 
            {
                var euro = egysor[j];
                //var productid = (codeStart + mainRow[j] + "." + mainColumn[i] + ".00");
                var productid = ("FSB.20." + mainRow[j] + "." + mainColumn[i] + ".00");
                checkId(productid);
                itemNoArray.push({productid, euro});
            }
        }
        console.log(itemNoArray);
    })   
}
//readMatrix('FSB.xlsx');
//readMatrix('SB.xlsx');

//Read Projcet file
console.log('Reading Project.xlsx');
function readProject(projectName) //3 PARAMETERES SB, FSB, PROJECTNAME
{
    readMatrix('FSB.xlsx');//SB + FSB PARAMETERS

    //workbook.xlsx.readFile('Project.xlsx')
    workbook.xlsx.readFile(projectName)
    .then(function() 
    {
        //Reads id and quantity coulumn from Project.xlsx.
        projectWorksheet = workbook.getWorksheet('Matten');
        projectProductId = projectWorksheet.getColumn('A').values;  
        quantity = projectWorksheet.getColumn('C').values;

        for (let i = 5; i < projectProductId.length; i++) 
        {
            var productid = projectProductId[i];
            checkId(productid);
            var value = quantity[i];
            idsAndQuantityArray.push({productid, value});
            //Stores id-s and quantities in an object.
        }
        console.log(idsAndQuantityArray);
        //return idsAndQuantityArray;

        //Looks for Project.xlsx id-s in the array based on FSB/SB.xlsx.
        for (let i = 0; i < idsAndQuantityArray.length; i++)
        {
            for (let j = 0; j < itemNoArray.length; j++)
            {
                if (idsAndQuantityArray[i].productid === itemNoArray[j].productid)
                {
                    let id = idsAndQuantityArray[i].productid;
                    let quantity = Number(idsAndQuantityArray[i].value);
                    let unitPrice = Number(itemNoArray[j].euro);
                    let quantityTimesunitPrice = quantity * unitPrice;
                    let fullPrice = Number(quantityTimesunitPrice.toFixed(3));
                    arrOfAllElements.push({id, quantity, unitPrice, fullPrice});
                }
            }
        }
        console.log(arrOfAllElements);

        //Writes out the sum of fullPrices. This is the final price that the client should pay.
        let sumFullPrice = 0;

        for (let i = 0; i < arrOfAllElements.length; i++)
        {
            sumFullPrice = sumFullPrice + arrOfAllElements[i].fullPrice;
        }
        let sum = sumFullPrice.toFixed(3);
        console.log("Total Preis: " + sum + " €");
    })

}
readProject('Project.xlsx');

//------------------------------------------------------------

/*console.log('Reading FSB.xlsx');
workbook.xlsx.readFile('FSB.xlsx')
    .then(function() 
    {
        worksheet = workbook.getWorksheet('Munka1');
        mainColumn = worksheet.getColumn('A').values;    
        mainRow = worksheet.getRow(2).values;

        for (let i = 2; i < mainColumn.length; i++) 
        {
            if (mainColumn[i] < 1000) 
            {
                mainColumn[i] = "0" + mainColumn[i];
            }
        }

        worksheet = workbook.getWorksheet('Munka1');
        for (let i = 3; i < mainColumn.length; i++) 
        {
            var egysor=worksheet.getRow(i).values;
            for (let j = 2; j < mainRow.length-1; j++) 
            {
                var euro = egysor[j];
                var productid = ("FSB.20."+ mainRow[j] + "." + mainColumn[i] + ".00");
                itemNoArray.push({productid, euro});
            }
        }
        console.log(itemNoArray);
        //return itemNoArray;
    })
    .then(function()
    {*/
    
//------------------------------------------------------------

/*let itemNoArray = 
    [{ productid: 'SB.20.1000.0150.00', euro: 10.88 },
    { productid: 'SB.20.1500.0150.00', euro: 13.22 },
    { productid: 'SB.20.2000.0150.00', euro: 15.47 },
    { productid: 'SB.20.2500.0150.00', euro: 17.88 },
    { productid: 'SB.20.3000.0150.00', euro: 20.05 }];

let idsAndQuantityArray = 
    [ { id: 'SB.20.2000.0150.00', value: '3' },
    { id: 'FSB.20.6000.0600.00', value: '3' },
    { id: 'FSB.20.3000.0600.00', value: '2' },
    { id: 'FSB.20.3000.0600.00', value: '1' },
    { id: 'FSB.20.3000.0600.00', value: '2' }];

let arrOfAllElements = [];*/

/*for (let i = 0; i < idsAndQuantityArray.length; i++)
{
    for (let j = 0; j < itemNoArray.length; j++)
    {
        if (idsAndQuantityArray[i].id === itemNoArray[j].productid)
        {
            let id = idsAndQuantityArray[i].id;
            let quantity = Number(idsAndQuantityArray[i].value);
            let unitPrice = Number(itemNoArray[j].euro);
            let quantityTimesunitPrice = quantity * unitPrice;
            let fullPrice = Number(quantityTimesunitPrice.toFixed(3));
            arrOfAllElements.push({id, quantity, unitPrice, fullPrice});
        }
    }
}
console.log(arrOfAllElements);*/

//------------------------------------------------------------

/*function checkId(id)
{
    if (id.match(/^[F]?SB.([0-9]{2}).([0-9]{4}).([0-9]{4}).([0-9]{2})$/))
    {
        console.log("true");
        return true;
    }
    else
    {
        console.log("false");
        return false;
    }
}
checkId("FSB.20.3000.0600.00");
checkId("SB.20.1500.0150.00");
checkId("S.20.1500.0150.00");*/



