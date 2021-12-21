/**
 * Genera un menú en la hoja de cálculo que permite generar los contratos para cada tipo
 */
function onOpen() {

    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('Crear');
    menu.addItem('CPSRPL', 'contratoDeServiciosconrpl');
    menu.addItem('CPS', 'contratoDePrestacion');
    menu.addItem('OS', 'ordendeServicios');
    ;
    menu.addToUi();


}

/**
 * Permite crear contratos de prestación de servicios con representante legal
 */
function contratoDeServiciosconrpl() {

    const googleTemplate = DriveApp.getFileById('1uhY4S4Do_Z3czBmgOW-khmTmGOB28dKor9iEfmBbLN8');
    const folder = DriveApp.getFolderById('11m3XK9Kd4MHRwsqrDrGDSVWL8_1yZBk0');
    const sheet1 = SpreadsheetApp.openById('1CgEI8rd4oBCBpUdfSUjW4j4i8zIlMhVZ8ggEF8sc5jc');
    const sheet = sheet1.getSheetByName('ARRAY');
    const range = sheet.getDataRange().getValues();

    range.forEach(function (range, index) {
        if (index === 0) return;
        if (range[59]) return;
        if (range[21] == '') return;
        if (range[21] == 'ORDEN DE SERVICIOS') return;
        if (range[21] == 'CONTRATO DE PRESTACIÓN DE SERVICIOS') return;
        if (range[29] == 'APROBADO GERENCIA') return;
        if (range[29] == 'APROBADO JURIDICA') return;
        const copy = googleTemplate.makeCopy(`Contrato de servicios ${range[4]}`, folder)
        const doc = DocumentApp.openById(copy.getId())
        const body = doc.getBody();

        body.replaceText('<< N de contrato >>', range[4]);
        body.replaceText('<<Fecha de inicio>>', range[15]);
        body.replaceText('<<Tipo de persona>>', range[2]);
        body.replaceText('<<Razon Social>>', range[5]);
        body.replaceText('<<Tipo de identificacion>>', range[1]);
        body.replaceText('<<N de identificacion>>', range[3]);
        body.replaceText('<<Nombre del representante>>', range[7]);
        body.replaceText('<<Tipo de identificacion del representante>>', range[8]);
        body.replaceText('<<N° de identificacion del representante>>', range[9]);
        body.replaceText('<<Ciudad del domicilio>>', range[10]);
        body.replaceText('<<Desarrollo del contrato>>', range[58]);
        body.replaceText('<<Valor hora en letras>>', range[47]);
        body.replaceText('<<Valor hora>>', range[46]);
        body.replaceText('<<Valor en letras>>', range[45]);
        body.replaceText('<<Valor sin IVA>>', range[23]);
        body.replaceText('<<Plazo meses>>', range[48]);
        body.replaceText('<<Dia inicio>>', range[49]);
        body.replaceText('<<Inicio en letras>>', range[55]);
        body.replaceText('<<año de inicio>>', range[51]);
        body.replaceText('<<Dia de terminacion>>', range[52]);
        body.replaceText('<<Terminacion en letras>>', range[56]);
        body.replaceText('<<año de terminación>>', range[54]);
        body.replaceText('<<Direccion del domicilio>>', range[11]);
        body.replaceText('<< Numero de telefono >>', range[13]);
        body.replaceText('<<Objeto>>', range[19]);
        body.replaceText('<<Alcance del contrato>>', range[18]);
        body.replaceText('<<Dia de elaboracion del contrato>>', range[57]);


        doc.saveAndClose();
        const url = doc.getUrl();
        sheet.getRange(index + 1, 60).setValue(url);


    })


}


/**
 * Permite crear contratos de prestación de servicios
 */
function contratoDePrestacion() {

    const googleTemplate = DriveApp.getFileById('1rYIbsK7pRN3dtpzHzK9dm0-GendBRguvoXSomFouCCs');
    const folder = DriveApp.getFolderById('11m3XK9Kd4MHRwsqrDrGDSVWL8_1yZBk0');
    const sheet1 = SpreadsheetApp.openById('1CgEI8rd4oBCBpUdfSUjW4j4i8zIlMhVZ8ggEF8sc5jc');
    const sheet = sheet1.getSheetByName('ARRAY');
    const range = sheet.getDataRange().getValues();

    range.forEach(function (range, index) {
        if (index === 0) return;
        if (range[59]) return;
        if (range[21] == '') return;
        if (range[21] == 'ORDEN DE SERVICIOS') return;
        if (range[21] == 'CONTRATO DE PRESTACIÓN DE SERVICIOS CON REPRESENTANTE LEGAL') return;
        if (range[29] == 'APROBADO GERENCIA') return;
        if (range[29] == 'APROBADO JURIDICA') return;
        const copy = googleTemplate.makeCopy(`Contrato de servicios ${range[4]}`, folder)
        const doc = DocumentApp.openById(copy.getId())
        const body = doc.getBody();

        body.replaceText('<< N de contrato >>', range[4]);
        body.replaceText('<<Fecha de inicio>>', range[15]);
        body.replaceText('<< Razon Social >>', range[5]);
        body.replaceText('<<Tipo de identificacion>>', range[1]);
        body.replaceText('<<N de identificacion>>', range[3]);
        body.replaceText('<<Ciudad del domicilio>>', range[10]);
        body.replaceText('<<Desarrollo del contrato>>', range[58]);
        body.replaceText('<<Valor hora en letras>>', range[47]);
        body.replaceText('<<Valor hora>>', range[46]);
        body.replaceText('<<Valor en letras>>', range[45]);
        body.replaceText('<<Valor sin IVA>>', range[23]);
        body.replaceText('<<Plazo meses>>', range[48]);
        body.replaceText('<<Dia inicio>>', range[49]);
        body.replaceText('<<Inicio en letras>>', range[55]);
        body.replaceText('<<año de inicio>>', range[51]);
        body.replaceText('<<Dia de terminacion>>', range[52]);
        body.replaceText('<<Terminacion en letras>>', range[56]);
        body.replaceText('<<año de terminación>>', range[54]);
        body.replaceText('<<Direccion del domicilio>>', range[11]);
        body.replaceText('<< Objeto >>', range[17]);
        body.replaceText('<< Numero de telefono >>', range[13]);
        body.replaceText('<<Dia de elaboracion del contrato>>', range[57]);

        doc.saveAndClose();
        const url = doc.getUrl();
        sheet.getRange(index + 1, 60).setValue(url);



    })



}

/**
 * Permite crear contratos de ordenes de servicios
 */
function ordendeServicios() {

    const googleTemplate = DriveApp.getFileById('1NLJzt7N2om3TMS6RuaWzfThKoYUw-k68sSHT88Mi818');
    const folder = DriveApp.getFolderById('1cbYpfqh2psib8klFyCBmQ6bF3F9ProSF');
    const sheet1 = SpreadsheetApp.openById('1CgEI8rd4oBCBpUdfSUjW4j4i8zIlMhVZ8ggEF8sc5jc');
    const sheet = sheet1.getSheetByName('ARRAY');
    const range = sheet.getDataRange().getValues();

    range.forEach(function (range, index) {
        if (index === 0) return;
        if (range[59]) return;
        if (range[21] == '') return;
        if (range[21] == 'CONTRATO DE PRESTACIÓN DE SERVICIOS') return;
        if (range[21] == 'CONTRATO DE PRESTACIÓN DE SERVICIOS CON REPRESENTANTE LEGAL') return;
        if (range[29] == 'APROBADO GERENCIA') return;
        if (range[29] == 'APROBADO JURIDICA') return;
        const copy = googleTemplate.makeCopy(`Orden de servicios ${range[4]}`, folder)
        const doc = DocumentApp.openById(copy.getId())
        const body = doc.getBody();

        body.replaceText('<<N de contrato>>', range[4]);
        body.replaceText('<<Fecha de inicio>>', range[15]);
        body.replaceText('<<Tipo de persona>>', range[2]);
        body.replaceText('<<Razon social>>', range[5]);
        body.replaceText('<<Tipo de identificacion>>', range[1]);
        body.replaceText('<<N de identificacion>>', range[3]);
        body.replaceText('<<Ciudad de domicilio>>', range[10]);
        body.replaceText('<<Desarrollo del contrato>>', range[58]);
        body.replaceText('<<Valor hora en letras>>', range[47]);
        body.replaceText('<<Valor hora>>', range[46]);
        body.replaceText('<<Valor en letras>>', range[45]);
        body.replaceText('<<Valor (sin IVA)>>', range[23]);
        body.replaceText('<<Plazo meses>>', range[48]);
        body.replaceText('<<Dia de inicio>>', range[49]);
        body.replaceText('<<Inicio en letras>>', range[55]);
        body.replaceText('<<año de inicio>>', range[51]);
        body.replaceText('<<Dia de terminacion>>', range[52]);
        body.replaceText('<<Terminacion en letras>>', range[56]);
        body.replaceText('<<año de terminación>>', range[54]);
        body.replaceText('<<Direccion del domicilio>>', range[11]);
        body.replaceText('<<Objeto>>', range[17]);
        body.replaceText('<<Alcance del contrato>>', range[18]);
        body.replaceText('<< Numero de telefono >>', range[13]);
        body.replaceText('<<Dia de elaboracion del contrato>>', range[57]);

        doc.saveAndClose();
        const url = doc.getUrl();
        sheet.getRange(index + 1, 60).setValue(url);



    })



}
