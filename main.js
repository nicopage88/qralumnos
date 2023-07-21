var fs = require('fs');
var pdf = require('dynamic-html-pdf');
var html = fs.readFileSync('template.html', 'utf8');
var qr = require('qr-image');

// Generate Excel to JSON
function genExcelToJson(){
    const excelToJson = require('convert-excel-to-json');
    const jsonResult = excelToJson({
        sourceFile: 'test.xlsx',
        columnToKey: {
            A: 'nombre',
            B: 'apellidos',
            C: 'curso'
        }
    });

    return jsonResult;
}

// Initialise PDF
function initPdf(){

    var jsonData = genExcelToJson(),
        finalArr = jsonData['Sheet 1'],
        currentObj,
        qrImgStr,
        qrLinStr,
        counter = 1;
    
    for(;counter<finalArr.length;counter++){   
        currentObj = finalArr[counter];
        // QR Generation
        qrLinStr = currentObj.name + '-' + currentObj.number + '-' + currentObj.email;     
        qrImgStr = qrImgGen(qrLinStr);
        // PDF Generate
        pdfGenInit(currentObj.name ,qrImgStr,qrLinStr);
    }
}

// QR Image Generation
function qrImgGen(qrLinStr){
    var qrImgStr = qr.imageSync(qrLinStr, { type: 'svg' });
    return qrImgStr;
}

// PDF Generate Initialise
function pdfGenInit(name,qrImgStr,qrLinStr){
    var pdfConfig = {
        type: 'file',     // 'file' or 'buffer'
        template: html,
        context: {
            data:{
                name: name,
                qrImg: qrImgStr,
                qrLink: qrLinStr
            }
        },
        path: "./output/output_" + name+".pdf"    // it is not required if type is buffer
    };
    crtPdf(pdfConfig);
}

// PDF Creation
function crtPdf(pdfConfig){
    var options = {
        timeout: 300000,
        format: "A4",
        orientation: "portrait",
        border: "10mm"
    };

    pdf.create(pdfConfig, options)
    .then(res => {
        console.log(res)
    })
    .catch(error => {
        console.error(error)
    });
}

initPdf();

