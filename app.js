//Read SB or FSB.

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
let fullPrice;
//Global variables.

//ID samples (write regex):
//SB20.2000.0150.00
//FSB.20.6000.0600.00

console.log('Reading SB.xlsx');
workbook.xlsx.readFile('SB.xlsx')
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
            var productid = ("SB.20."+ mainRow[j] + "." + mainColumn[i] + ".00");
            //Use variable instead of static value above.
            itemNoArray.push({productid, euro});
            }
        }
        console.log(itemNoArray);
        //return itemNoArray;
    })
    .then(function()
    {
    
//------------------------------------------------------------

    console.log('Reading Project.xlsx');
    workbook.xlsx.readFile('Project.xlsx')
        .then(function() 
        {
            //Reads id and quantity coulumn from Project.xlsx.
            projectWorksheet = workbook.getWorksheet('Matten');
            projectProductId = projectWorksheet.getColumn('A').values;  
            quantity = projectWorksheet.getColumn('C').values;

            for (let i = 5; i < projectProductId.length; i++) 
            {
                var id = projectProductId[i];
                var value = quantity[i];
                idsAndQuantityArray.push({id, value});
                //Stores id-s and quantities in an object.
            }
            console.log(idsAndQuantityArray);
            //return idsAndQuantityArray;

            //Looks for Project.xlsx id-s in the array based on FSB/SB.xlsx.
            console.log('Final array with all data');
            for (let i = 0; i < idsAndQuantityArray.length; i++)
            {
                for (let j = 0; j < itemNoArray.length; j++)
                {
                    if (idsAndQuantityArray[i].id === itemNoArray[j].productid)
                    {
                        let id = idsAndQuantityArray[i].id;
                        let quantity = Number(idsAndQuantityArray[i].value);
                        let unitPrice = Number(itemNoArray[j].euro);
                        let quantityTimesunitPrice = quantity * unitPrice;
                        //let fullPrice = Number(quantityTimesunitPrice.toFixed(3));
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
            console.log(sum);
        })
    })

//------------------------------------------------------------

/*.then(function()
{
    var worksheet = workbook.addWorksheet('Munka1');//WORKSHEET ATNEVEZ

    worksheet.getColumn('A').values = fullPrice;

    workbook.xlsx.writeFile("Preisangebot.xlsx")
    .then(function() 
    {
        console.log(worksheet.getRow('A').values);
    });
})
//New excel generator.*/

//var Excel = require("exceljs");
//var workbook = new Excel.Workbook();
.then(function()
{
            let valuesA;
            let valuesB;
            let valuesC;
            let valuesD;
            let valuesE;
            let valuesF;
            let valuesG;
            let valuesH;

        workbook.xlsx.readFile('Project.xlsx')
        //workbook.xlsx.readFile('SB.xlsx')
        .then(function() 
        {
        //Reads all the columns of the Project.xlsx.
            /*projectWorksheet.getColumn('A').numFmt='#.00';
            projectWorksheet.getColumn('B').numFmt='#.00';
            projectWorksheet.getColumn('C').numFmt='#.00';
            projectWorksheet.getColumn('D').numFmt='#.00';
            projectWorksheet.getColumn('E').numFmt='#.00';
            projectWorksheet.getColumn('F').numFmt='#.00';
            projectWorksheet.getColumn('G').numFmt='#.00';
            projectWorksheet.getColumn('H').numFmt='#.00';*/

            valuesA = projectWorksheet.getColumn('A').values;
            valuesB = projectWorksheet.getColumn('B').values;
            valuesC = projectWorksheet.getColumn('C').values;
            valuesD = projectWorksheet.getColumn('D').values;
            valuesE = projectWorksheet.getColumn('E').values;
            valuesF = projectWorksheet.getColumn('F').values;
            valuesG = projectWorksheet.getColumn('G').values;
            valuesH = projectWorksheet.getColumn('H').values;

            console.log("-----------------------------------------");
            for(var i=0;i<valuesG.length;i++){
                console.log(typeof valuesG[i] + " " + valuesG[i]);
            }
            //Empty cells type is undefined - all others are strings.

            /*console.log(valuesA);
            console.log(valuesB);
            console.log(valuesC);
            console.log(valuesD);
            console.log(valuesE);
            console.log(valuesF);
            console.log(valuesG);
            console.log(valuesH);*/
        })
        .then(function() 
        {
    });
})

//------------------------------------------------------------
//------------------------------------------------------------

//Reads from file:
//https://www.npmjs.com/package/exceljs

//----------------

//get cell

/*var Excel = require("exceljs");
var workbook = new Excel.Workbook();
workbook.xlsx.readFile('Arlista.xlsx')
    .then(function() {
       console.log(workbook.getWorksheet("Munka1").getRow(4).getCell(1).value);
    });
	
//----------------
 
//get whole table

var Excel = require("exceljs");
var workbook = new Excel.Workbook();
workbook.xlsx.readFile('Arlista.xlsx')
    .then(function() {
        worksheet = workbook.getWorksheet('Munka1');
        worksheet.eachRow({ includeEmpty: true },function(row, rowNumber) {
          console.log("Row " + rowNumber + " = " + JSON.stringify(row.values));
        });
     });

//----------------

//get row

var Excel = require("exceljs");
var workbook = new Excel.Workbook();
workbook.xlsx.readFile('Arlista.xlsx')
    .then(function() {
        worksheet = workbook.getWorksheet('Munka1');
        var row = worksheet.getRow(5).values;
        console.log(row);
    });
	

//----------------

//get column

var Excel = require("exceljs");
var workbook = new Excel.Workbook();
workbook.xlsx.readFile('Arlista.xlsx')
    .then(function() {
        worksheet = workbook.getWorksheet('Munka1');
        var row = worksheet.getColumn('B').values;
        console.log(row);
    });*/

//------------------------------------------------------------
//------------------------------------------------------------

//Writes file:
//https://www.npmjs.com/package/exceljs

//----------------	

//creates new table and worksheet ('Work')

/*var Excel = require("exceljs");
var workbook = new Excel.Workbook();
var sheet = workbook.addWorksheet('Work');

workbook.xlsx.writeFile("PriceList.xlsx")
    .then(function() {
        // console.log("xls file is written.");
    });	

//----------------	

//writes header

var Excel = require("exceljs");
var workbook = new Excel.Workbook();
var worksheet = workbook.addWorksheet('Work');

worksheet.columns = [
    { header: 'Word', key: 'word', width: 36 },
    { header: 'Definition', key: 'def', width: 120 }
];

workbook.xlsx.writeFile("PriceList.xlsx")
    .then(function() {
        console.log(worksheet.getRow(1).values);
    });

//----------------

//writes column

var Excel = require("exceljs");
var workbook = new Excel.Workbook();
var worksheet = workbook.addWorksheet('Work');

worksheet.getColumn('A').values = [1,2,3,4,5];

workbook.xlsx.writeFile("PriceList.xlsx")
    .then(function() {
        console.log(worksheet.getColumn('A').values);
    });

//----------------	

//writes cell
	
var Excel = require("exceljs");
var workbook = new Excel.Workbook();
var worksheet = workbook.addWorksheet('Work');

worksheet.getCell('C3').value = "alma";

workbook.xlsx.writeFile("PriceList.xlsx")
    .then(function() {
        console.log(worksheet.getCell('C3').value);
    });
	
//----------------
	
//writes multiple rows
	
var Excel = require("exceljs");
var workbook = new Excel.Workbook();
var worksheet = workbook.addWorksheet('Work');

worksheet.getRow(1).values = ["A","B", "C"];
worksheet.getRow(2).values = [1,"a", "2"];
worksheet.getRow(3).values = [2,"b", "3"];

workbook.xlsx.writeFile("PriceList.xlsx")
    .then(function() {
        console.log(worksheet.getRow(1).values);
        console.log(worksheet.getRow(2).values);
        console.log(worksheet.getRow(3).values);
    });*/

//------------------------------------------------------------
//------------------------------------------------------------

/*let sumFullPrice;

for (let i = 0; i < arrOfAllElements.length; i++)
{
    sumFullPrice = sumFullPrice + arrOfAllElements[i].fullPrice;
}*/


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

//-----------------------

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

//-----------------------

//workbook.xlsx.readFile('FSB.xlsx')

/*let idStart;

if(readfile == 'FSB.xlsx')
{
    idStart = "FSB.20.";
}
else if (readfile == 'SB.xlsx')
{
    idStart = "SB.20.";
}*/



