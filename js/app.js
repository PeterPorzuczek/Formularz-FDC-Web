var underscore = _;

var columns = [ "A", "B", "C", "D", "E", 
                "F", "G", "H", "I", "J", 
                "K", "L", "M", "N", "O", 
                "P", "Q", "R", "S", "T", 
                "U", "V", "W", "X", "Y", "Z", 
                "AA", "AB", "AC", "AD", 
                "AE", "AF", "AG", "AH" ];

var containers = {
    "head": "head-container",
    "relativesCount": "relatives-count-container",
    "generatedContent": "generated-form-container",
    "map": "map-container"
};

var forms = {
    "relativesCountForm": "inputRelativesCountForm",
    "generationForm": "generatedForm"
};

var headSection = {
    "title": "Formularz FDC",
    "description": "Strona udostępnia formularz, który może być wykorzystany w przedsiębiorstwie celem ułatwienia pracownikom zgłoszenia członków rodziny do ubezpieczenia zdrowotnego. Narzędzie dostarczane za pośrednictwem strony pozwala wypełnić formularz wewnątrz przeglądarki i zapisać go w pliku xlsx na komputerze użytkownika. Strona nie przechowuje danych. Informacje wprowadzane przez użytkownika występują wyłącznie w przeglądarce."
};

var yesNoOptions = {
    "yes": { name: "Tak", value: "tak" },
    "no": { name: "Nie", value: "nie" }
};
var idOptions = [   { name: "Pesel", value: "pesel", id: "" }, 
                    { name: "Dowód osobisty", value: "dow. os.", id: "1" }, 
                    { name: "Paszport", value: "paszport", id: "2" }    
                ];  

var relationshipOptions = [ { name: "Małżonek", code: "01" }, 
                            { name: "Dziecko własne, przysposobione lub dziecko małżonka", code: "11" }, 
                            { name: "Wnuk albo dziecko obce, dla którego ustanowiono opiekę, albo dziecko obce w ramach rodziny zastępczej lub rodzinnego domu dziecka", code: "21" }, 
                            { name: "Matka", code: "30" }, 
                            { name: "Ojciec", code: "31" }, 
                            { name: "Macocha", code: "32" }, 
                            { name: "Ojczym", code: "33" }, 
                            { name: "Babka", code: "40" }, 
                            { name: "Dziadek", code: "41" }, 
                            { name: "Osoby przysposabiające osoby ubezpieczone", code: "50" }, 
                            { name: "Inni wstępni pozostający z ubezpieczonym we współnym gospodarstwie domowym", code: "60" }  
                          ];
var disabilityOptions = [   { name: "Nie posiada orzeczenia o niepełnosprawności", code: "0" }, 
                            { name: "Posiada orzeczenie o lekkim stopniu niepełnosprawności", code: "1" }, 
                            { name: "Posiada orzeczenie o umiarkowanym stopniu niepełnosprawności", code: "2" }, 
                            { name: "Posiada orzeczenie o znacznym stopniu niepełnosprawności", code: "3" }, 
                            { name: "Posiada orzeczenie o niepełnosprawności wydawane osobom do 16 roku życia", code: "4" }     
                        ];

var voivodeshipOptions = [  { name: "Mazowieckie", value: "mazowieckie" }, 
                            { name: "Śląskie", value: "śląskie" }, 
                            { name: "Wielkopolskie", value: "wielkopolskie" }, 
                            { name: "Małopolskie", value: "małopolskie" }, 
                            { name: "Dolnośląskie", value: "dolnośląskie" }, 
                            { name: "Łódzkie", value: "łódzkie" }, 
                            { name: "Pomorskie", value: "pomorskie" }, 
                            { name: "Lubelskie", value: "lubelskie" }, 
                            { name: "Podkarpackie", value: "podkarpackie" }, 
                            { name: "Kujawsko-pomorskie", value: "kujawsko-pomorskie" }, 
                            { name: "Zachodniopomorskie", value: "zachodniopomorskie" }, 
                            { name: "Warmińsko-mazurskie", value: "warmińsko-mazurskie" }, 
                            { name: "Świętokrzyskie", value: "świętokrzyskie" }, 
                            { name: "Podlaskie", value: "podlaskie" }, 
                            { name: "Lubuskie", value: "lubuskie" }, 
                            { name: "Opolskie", value: "opolskie" }     
                        ];
var formGroupIdPrefix = {
    "relativesCountSectionTitle": "relativesCountSectionTitleFormGroupID",
    "identificationSectionTitle": "identificationSectionTitleFormGroupID",
    "relativeIdentificationSectionTitle": "relativeIdentificationSectionTitleFormGroupID",
    "relativeAddressSectionTitle": "relativeAddressSectionTitleFormGroupID",
    "locationSectionTitle": "locationSectionTitleFormGroupID",
    "generationSectionTitle": "generationSectionTitleFormGroupID",
    "employeeID": "employeeIDTextboxFormGroupID",
    "employeeName": "employeeNameTextboxFormGroupID",
    "employeeSurname": "employeeSurnameTextboxFormGroupID",
    "relativeName": "relativeNameTextboxFormGroupID",
    "relativeSurname": "relativeSurnameTextboxFormGroupID",
    "relativeBirth": "relativeBirthTextboxFormGroupID",
    "relativeCommonGround": "relativeCommonGroundTextboxFormGroupID",
    "relativePesel": "relativePeselTextboxFormGroupID",
    "relativeDocument1": "relativeDocument1TextboxFormGroupID",
    "relativeDocument2": "relativeDocument2TextboxFormGroupID",
    "relativeDocumentID": "relativeDocumentIDTextboxFormGroupID",
    "relativeRelationship": "relativeRelationshipTextboxFormGroupID",
    "relativeDisability": "relativeDisabilityTextboxFormGroupID",
    "relativeVoivodeship": "relativeVoivodeshipTextboxFormGroupID",
    "relativeCommunity": "relativeCommunityTextboxFormGroupID",
    "relativePlace": "relativePlaceTextboxFormGroupID",
    "relativePostal": "relativePostalTextboxFormGroupID",
    "relativeStreet": "relativeStreetTextboxFormGroupID",
    "relativeFlatNumber": "relativeFlatNumberTextboxFormGroupID",
    "relativeApartmentNumber": "relativeApartmentNumberTextboxFormGroupID",
    "relativeTelephoneNumber": "relativeTelephoneNumberTextboxFormGroupID",
    "spreadsheetLocation": "spreadsheetLocationTextboxFormGroupID"
};
var textboxIdPrefix = {
    "employeeID": "employeeIDTextboxID",
    "employeeName": "employeeNameTextboxID",
    "employeeSurname": "employeeSurnameTextboxID",
    "relativeName": "relativeNameTextboxID",
    "relativeSurname": "relativeSurnameTextboxID",
    "relativeBirth": "relativeBirthTextboxID",
    "relativeCommonGround": "relativeCommonGroundTextboxID",
    "relativePesel": "relativePeselTextboxID",
    "relativeDocument1": "relativeDocument1TextboxID",
    "relativeDocument2": "relativeDocument2TextboxID",
    "relativeDocumentID": "relativeDocumentIDTextboxID",
    "relativeRelationship": "relativeRelationshipTextboxID",
    "relativeDisability": "relativeDisabilityTextboxID",
    "relativeVoivodeship": "relativeVoivodeshipTextboxID",
    "relativeCommunity": "relativeCommunityTextboxID",
    "relativePlace": "relativePlaceTextboxID",
    "relativePostal": "relativePostalTextboxID",
    "relativeStreet": "relativeStreetTextboxID",
    "relativeFlatNumber": "relativeFlatNumberTextboxID",
    "relativeApartmentNumber": "relativeApartmentNumberTextboxID",
    "relativeTelephoneNumber": "relativeTelephoneNumberTextboxID",
    "spreadsheetLocation": "spreadsheetLocationTextboxID",
    "relativesCountSection": "relativesCountSectionTextboxID"
};

var relativesCountSection = function relativesCountSection(num) {
    return [{
                "value": "Rozpocznij",
                "type": "head",
                "formId": "relativesCountSectionTitleFormGroupID" + num,
                "num": num
            }, {
                "label": "Liczba członków rodziny zgłaszanych do ubezpieczenia zdrowotnego:",
                "prefix": textboxIdPrefix.relativesCountSection + num,
                "id": textboxIdPrefix.relativesCountSection + num,
                "name": "relativesCountSectionTextboxNumber" + num,
                "formId": "relativesCountSectionTextboxFormGroupID" + num,
                "type": "number",
                "minlength": 1,
                "maxlength": 2,
                "required": true,
                "date": false,
                "digits": true,
                "min": 0,
                "max": 21,
                "num": num
            }, {
                "label": "Dalej",
                "type": "submit",
                "class": "btn btn-primary w-100"
            }];
};

var employeeIdentificationSection = function employeeIdentificationSection(num) {
    return [{
                "value": num + ". DANE IDENTYFIKACYJNE OSOBY UBEZPIECZONEJ (PRACOWNIKA)",
                "type": "head",
                "class": "",
                "formId": formGroupIdPrefix.identificationSectionTitle + num,
                "num": num
            }, {
                "label": "Numer kontrolny pracownika:",
                "prefix": textboxIdPrefix.employeeID,
                "id": textboxIdPrefix.employeeID + num,
                "name": "employeeIDTextboxNumber" + num,
                "formId": formGroupIdPrefix.employeeID + num,
                "type": "number",
                "minlength": 5,
                "maxlength": 5,
                "required": true,
                "date": false,
                "digits": true,
                "num": num
            }, {
                "label": "Imię pracownika:",
                "prefix": textboxIdPrefix.employeeName,
                "id": textboxIdPrefix.employeeName + num,
                "name": "employeeNameTextboxNumber" + num,
                "formId": formGroupIdPrefix.employeeName + num,
                "type": "text",
                "minlength": 0,
                "maxlength": 29,
                "required": true,
                "date": false,
                "digits": false,
                "num": num
            }, {
                "label": "Nazwisko pracownika:",
                "prefix": textboxIdPrefix.employeeSurname,
                "id": textboxIdPrefix.employeeSurname + num,
                "name": "employeeSurnameTextboxNumber" + num,
                "formId": formGroupIdPrefix.employeeSurname + num,
                "type": "text",
                "minlength": 0,
                "maxlength": 58,
                "required": true,
                "date": false,
                "digits": false,
                "num": num
            }];
};

var relativeIdentificationSection = function relativeIdentificationSection(num) {
    return [{
                "value": num + ". DANE O CZŁONKU RODZINY, UPRAWNIONYM DO ŚWIADCZEŃ Z UBEZPIECZENIA ZDROWOTNEGO",
                "class": "",
                "type": "head",
                "formId": formGroupIdPrefix.relativeIdentificationSectionTitle + num
            }, {
                "label": "Imię członka rodziny:",
                "prefix": textboxIdPrefix.relativeName,
                "id": textboxIdPrefix.relativeName + num,
                "name": "relativeNameTextboxNumber" + num,
                "formId": formGroupIdPrefix.relativeName + num,
                "type": "text",
                "minlength": 0,
                "maxlength": 29,
                "required": true,
                "date": false,
                "digits": false,
                "num": num
            }, {
                "label": "Nazwisko członka rodziny:",
                "prefix": textboxIdPrefix.relativeSurname,
                "id": textboxIdPrefix.relativeSurname + num,
                "name": "relativeSurnameTextboxNumber" + num,
                "formId": formGroupIdPrefix.relativeSurname + num,
                "type": "text",
                "minlength": 0,
                "maxlength": 29,
                "required": true,
                "date": false,
                "digits": false,
                "num": num
            }, {
                "label": "Data urodzenia członka rodziny: (dd-mm-rrrr)",
                "prefix": textboxIdPrefix.relativeBirth,
                "id": textboxIdPrefix.relativeBirth + num,
                "name": "relativeBirthTextboxNumber" + num,
                "formId": formGroupIdPrefix.relativeBirth + num,
                "type": "text",
                "minlength": 10,
                "maxlength": 10,
                "required": true,
                "date": true,
                "digits": false,
                "num": num
            }, {
                "label": "Pozostaje we wspólnym gospodarstwie domowym z pracownikiem:",
                "prefix": textboxIdPrefix.relativeCommonGround,
                "id": textboxIdPrefix.relativeCommonGround + num,
                "name": "relativeCommonGroundTextboxNumber" + num,
                "formId": formGroupIdPrefix.relativeCommonGround + num,
                "required": true,
                "choices": [["", ""], [yesNoOptions.yes.value, yesNoOptions.yes.name], [yesNoOptions.no.value, yesNoOptions.no.name]],
                "type": "select",
                "num": num
            }, {
                "label": "Czy członek rodziny posiada pesel:",
                "prefix": textboxIdPrefix.relativePesel,
                "id": textboxIdPrefix.relativePesel + num,
                "name": "relativePeselTextboxNumber" + num,
                "formId": formGroupIdPrefix.relativePesel + num,
                "required": true,
                "choices": [["", ""], [yesNoOptions.yes.value, yesNoOptions.yes.name], [yesNoOptions.no.value, yesNoOptions.no.name]],
                "type": "select",
                "num": num
            }, {
                "label": "Rodzaj identyfikacji:",
                "prefix": textboxIdPrefix.relativeDocument1,
                "id": textboxIdPrefix.relativeDocument1 + num,
                "name": "relativeDocument1TextboxNumber" + num,
                "formId": formGroupIdPrefix.relativeDocument1 + num,
                "required": true,
                "choices": [[idOptions[0].value, idOptions[0].name]],
                "type": "select",
                "num": num
            }, {
                "label": "Rodzaj identyfikacji:",
                "prefix": textboxIdPrefix.relativeDocument2,
                "id": textboxIdPrefix.relativeDocument2 + num,
                "name": "relativeDocument2TextboxNumber" + num,
                "formId": formGroupIdPrefix.relativeDocument2 + num,
                "required": true,
                "choices": [["", ""], [idOptions[1].value, idOptions[1].name], [idOptions[2].value, idOptions[2].name]],
                "type": "select",
                "num": num
            }, {
                "label": "Pesel / seria i nr dow. os / seria i nr paszportu członka rodziny:",
                "prefix": textboxIdPrefix.relativeDocumentID,
                "id": textboxIdPrefix.relativeDocumentID + num,
                "name": "relativeDocumentIDTextboxNumber" + num,
                "formId": formGroupIdPrefix.relativeDocumentID + num,
                "type": "text",
                "minlength": 0,
                "maxlength": 11,
                "required": true,
                "date": false,
                "digits": false,
                "num": num
            }, {
                "label": "Pokrewienstwo członka rodziny:",
                "prefix": textboxIdPrefix.relativeRelationship,
                "id": textboxIdPrefix.relativeRelationship + num,
                "name": "relativeRelationshipTextboxNumber" + num,
                "formId": formGroupIdPrefix.relativeRelationship + num,
                "required": true,
                "choices": [["", ""], [relationshipOptions[0].name, relationshipOptions[0].name], [relationshipOptions[1].name, relationshipOptions[1].name], [relationshipOptions[2].name, relationshipOptions[2].name], [relationshipOptions[3].name, relationshipOptions[3].name], [relationshipOptions[4].name, relationshipOptions[4].name], [relationshipOptions[5].name, relationshipOptions[5].name], [relationshipOptions[6].name, relationshipOptions[6].name], [relationshipOptions[7].name, relationshipOptions[7].name], [relationshipOptions[8].name, relationshipOptions[8].name], [relationshipOptions[9].name, relationshipOptions[9].name], [relationshipOptions[10].name, relationshipOptions[10].name]],
                "type": "select",
                "num": num
            }, {
                "label": "Niepełnosprawność członka rodziny:",
                "prefix": textboxIdPrefix.relativeDisability,
                "id": textboxIdPrefix.relativeDisability + num,
                "name": "relativeDisabilityTextboxNumber" + num,
                "formId": formGroupIdPrefix.relativeDisability + num,
                "required": true,
                "choices": [["", ""], [disabilityOptions[0].name, disabilityOptions[0].name], [disabilityOptions[1].name, disabilityOptions[1].name], [disabilityOptions[2].name, disabilityOptions[2].name], [disabilityOptions[3].name, disabilityOptions[3].name], [disabilityOptions[4].name, disabilityOptions[4].name]],
                "type": "select",
                "num": num
            }];
};

var relativeAddressSection = function relativeAddressSection(num) {
    return [{
                "value": num + ".1. ADRES ZAMIESZKANIA CZŁONKA RODZINY",
                "valueAdd": "(wpisać, jeśli adres zamieszkania zgłaszanego członka rodziny jest inny, niż adres zamieszkania ubezpieczonego)",
                "class": "",
                "classAdd": "error",
                "type": "head2",
                "formId": formGroupIdPrefix.relativeAddressSectionTitle + num,
                "num": num
            }, {
                "label": "Województwo członka rodziny:",
                "prefix": textboxIdPrefix.relativeVoivodeship,
                "id": textboxIdPrefix.relativeVoivodeship + num,
                "name": "relativeVoivodeshipTextboxNumber" + num,
                "formId": formGroupIdPrefix.relativeVoivodeship + num,
                "required": true,
                "choices": [["", ""], [voivodeshipOptions[0].value, voivodeshipOptions[0].name], [voivodeshipOptions[1].value, voivodeshipOptions[1].name], [voivodeshipOptions[2].value, voivodeshipOptions[2].name], [voivodeshipOptions[3].value, voivodeshipOptions[3].name], [voivodeshipOptions[4].value, voivodeshipOptions[4].name], [voivodeshipOptions[5].value, voivodeshipOptions[5].name], [voivodeshipOptions[6].value, voivodeshipOptions[6].name], [voivodeshipOptions[7].value, voivodeshipOptions[7].name], [voivodeshipOptions[8].value, voivodeshipOptions[8].name], [voivodeshipOptions[9].value, voivodeshipOptions[9].name], [voivodeshipOptions[10].value, voivodeshipOptions[10].name], [voivodeshipOptions[11].value, voivodeshipOptions[11].name], [voivodeshipOptions[12].value, voivodeshipOptions[12].name], [voivodeshipOptions[13].value, voivodeshipOptions[13].name], [voivodeshipOptions[14].value, voivodeshipOptions[14].name], [voivodeshipOptions[15].value, voivodeshipOptions[15].name]],
                "type": "select",
                "num": num
            }, {
                "label": "Gmina członka rodziny:",
                "prefix": textboxIdPrefix.relativeCommunity,
                "id": textboxIdPrefix.relativeCommunity + num,
                "name": "relativeCommunityTextboxNumber" + num,
                "formId": formGroupIdPrefix.relativeCommunity + num,
                "type": "text",
                "minlength": 0,
                "maxlength": 29,
                "required": true,
                "date": false,
                "digits": false,
                "num": num
            }, {
                "label": "Miejscowość członka rodziny:",
                "prefix": textboxIdPrefix.relativePlace,
                "id": textboxIdPrefix.relativePlace + num,
                "name": "relativePlaceTextboxNumber" + num,
                "formId": formGroupIdPrefix.relativePlace + num,
                "type": "text",
                "minlength": 0,
                "maxlength": 22,
                "required": true,
                "date": false,
                "digits": false,
                "num": num
            }, {
                "label": "Kod pocztowy członka rodziny:",
                "prefix": textboxIdPrefix.relativePostal,
                "id": textboxIdPrefix.relativePostal + num,
                "name": "relativePostalTextboxNumber" + num,
                "formId": formGroupIdPrefix.relativePostal + num,
                "type": "text",
                "minlength": 6,
                "maxlength": 6,
                "required": true,
                "date": false,
                "digits": false,
                "num": num
            }, {
                "label": "Ulica członka rodziny:",
                "prefix": textboxIdPrefix.relativeStreet,
                "id": textboxIdPrefix.relativeStreet + num,
                "name": "relativeStreetTextboxNumber" + num,
                "formId": formGroupIdPrefix.relativeStreet + num,
                "type": "text",
                "minlength": 0,
                "maxlength": 29,
                "required": true,
                "date": false,
                "digits": false,
                "num": num
            }, {
                "label": "Numer domu członka rodziny:",
                "prefix": textboxIdPrefix.relativeFlatNumber,
                "id": textboxIdPrefix.relativeFlatNumber + num,
                "name": "relativeFlatNumberTextboxNumber" + num,
                "formId": formGroupIdPrefix.relativeFlatNumber + num,
                "type": "text",
                "minlength": 0,
                "maxlength": 8,
                "required": true,
                "date": false,
                "digits": false,
                "num": num
            }, {
                "label": "Numer lokalu członka rodziny:",
                "prefix": textboxIdPrefix.relativeApartmentNumber,
                "id": textboxIdPrefix.relativeApartmentNumber + num,
                "name": "relativeApartmentNumberTextboxNumber" + num,
                "formId": formGroupIdPrefix.relativeApartmentNumber + num,
                "type": "text",
                "minlength": 0,
                "maxlength": 5,
                "required": true,
                "date": false,
                "digits": false,
                "num": num
            }, {
                "label": "Numer telefonu członka rodziny:",
                "prefix": textboxIdPrefix.relativeTelephoneNumber,
                "id": textboxIdPrefix.relativeTelephoneNumber + num,
                "name": "relativeTelephoneNumberTextboxNumber" + num,
                "formId": formGroupIdPrefix.relativeTelephoneNumber + num,
                "type": "text",
                "minlength": 0,
                "maxlength": 14,
                "required": false,
                "date": false,
                "digits": false,
                "num": num
            }];
};

var locationIdentificationSection = function locationIdentificationSection(num) {
    return [{
                "value": num + ". OŚWIADCZENIE OSOBY ZGŁASZAJĄCEJ DANE",
                "class": "",
                "type": "head",
                "formId": formGroupIdPrefix.locationSectionTitle + num,
                "num": num
            }, {
                "label": "Miejsce oświadczenia (miejscowość):",
                "prefix": textboxIdPrefix.spreadsheetLocation,
                "id": "spreadsheetLocationTextboxID" + num,
                "name": "spreadsheetLocationNumber" + num,
                "formId": formGroupIdPrefix.spreadsheetLocation + num,
                "type": "text",
                "minlength": 0,
                "maxlength": 29,
                "required": true,
                "date": false,
                "digits": false,
                "num": num
            }];
};

var generationSection = function generationSection(num) {
    return [{
                "value": "Wygeneruj:",
                "class": "",
                "type": "head",
                "formId": formGroupIdPrefix.generationSectionTitle + num,
                "num": num
            }, {
                "label": "Generuj z danymi",
                "id": "spreadsheetWithDataButton" + num,
                "type": "submit",
                "class": "btn btn-primary w-100 h-5 my-3",
                "num": num
            }, {
                "label": "Generuj pusty",
                "id": "spreadsheetEmptyButton" + num,
                "type": "button",
                "class": "btn btn-primary w-100 h-5 my-2",
                "num": num
            }];
};

var xlsxAddress = function xlsxAddress(n) {
    return "assets/SourceXLSX/FormularzPusty-" + n + ".xlsx";
};

var Promise = XlsxPopulate.Promise,
    xlsxOrgResponse,
    urlInput,
    relativesCount;

document.getElementById("textTitle").innerHTML = headSection.title;
document.getElementById("textDescription").innerHTML = headSection.description;

$(window).bind("pageshow", function () {
    for (var key in forms) {
        if (forms.hasOwnProperty(key)) {
            var form = $('#' + forms[key]);

            form[0].reset();
        }
    }
    setContainersDisplayState("block", "block", "none");
});

function start() {
    setForm(forms.relativesCountForm, relativesCountSection(0), setRelatives);
};

function setForm(id, sectionFields, action) {
    var formId = id,
        fields = sectionFields;

    var form = FormForm($('#' + formId), fields);
    var rules = rulesMaker(fields);

    form.render();

    $("#" + formId).validate({
        rules: rules
    });

    $("#" + formId).submit(function (e) {
        e.preventDefault();
        if ($('#' + formId).valid()) {

            action(fields);

            return false;
        }
    });
}

function setRelatives(fields) {
    var relativesFieldId = fields[1].id;
    relativesCount = document.getElementById(relativesFieldId).value;
    urlInput = xlsxAddress(relativesCount);

    getRelativesWorkbook(generateForm);
}

function generateForm() {
    setContainersDisplayState("none", "none", "block");

    var formArray = concatForm();

    setForm(forms.generationForm, formArray, getSpreadsheetWithData);

    formPostGenerationPreparation(formArray);
}

function getSpreadsheetWithData(formArray) {
    var data = dataGet(formArray);
    generateEmptyBlob(false, data);
}

function dataGet(formArray) {
    var employeeIdentification = [],
        relativeIdentification = [],
        locationIdentification = [];

    employeeIdentification.push(employeeIdentificationDataGet(formArray, 1));

    var sectionsCount = parseInt(relativesCount) + 2;
    for (var i = 2; i < sectionsCount; i++) {
        relativeIdentification.push(relativeIdentificationDataGet(formArray, i));
    }

    locationIdentification.push(locationIdentificationDataGet(formArray, sectionsCount));

    return { "employeeIdentification": employeeIdentification,
             "relativeIdentification": relativeIdentification,
             "locationIdentification": locationIdentification };
}

function EmployeeIdentification() {
    var date, 
        relativesCount, 
        id, 
        name, 
        surname;
}

function RelativeIdentification() {
    var name, 
        surname, 
        peselCheck, 
        pesel, 
        birth, 
        commonGround, 
        idNum, 
        idType, 
        idId, 
        relationship, 
        relationshipId, 
        disability, 
        disabilityId, 
        voivodeship, 
        community, 
        place, 
        postal, 
        street, 
        flatNum, 
        apartmentNum, 
        telephone;
}
function LocationIdentification() {
    var date, 
        place;
}

function employeeIdentificationDataGet(formArray, num) {
    var employeeId                  = new EmployeeIdentification();
        employeeId.date             = dateToDDMMYYYY(new Date());
        employeeId.relativesCount   = relativesCount;
        employeeId.id               = valueById(underscore.where(formArray, {   "prefix": textboxIdPrefix.employeeID,
                                                                                "num": num })[0].id);
        employeeId.name             = valueById(underscore.where(formArray, {   "prefix": textboxIdPrefix.employeeName,
                                                                                "num": num })[0].id);
        employeeId.surname          = valueById(underscore.where(formArray, {   "prefix": textboxIdPrefix.employeeSurname,
                                                                                "num": num })[0].id);
    return employeeId;
}

function locationIdentificationDataGet(formArray, num) {
    var locationId                  = new LocationIdentification();
        locationId.date             = dateToDDMMYYYY(new Date());
        locationId.place            = valueById(underscore.where(formArray, {   "prefix": textboxIdPrefix.spreadsheetLocation,
                                                                                "num": num })[0].id);
    return locationId;
}

function relativeIdentificationDataGet(formArray, num) {
    var peselTemp;
    var relativeId                  = new RelativeIdentification();
        relativeId.name             = valueById(underscore.where(formArray, {   "prefix": textboxIdPrefix.relativeName,
                                                                                "num": num })[0].id);
        relativeId.surname          = valueById(underscore.where(formArray, {   "prefix": textboxIdPrefix.relativeSurname,
                                                                                "num": num })[0].id);
        relativeId.birth            = valueById(underscore.where(formArray, {   "prefix": textboxIdPrefix.relativeBirth,
                                                                                "num": num })[0].id);
        relativeId.commonGround     = valueById(underscore.where(formArray, {   "prefix": textboxIdPrefix.relativeCommonGround,
                                                                                "num": num })[0].id);
        relativeId.peselCheck       = valueById(underscore.where(formArray, {   "prefix": textboxIdPrefix.relativePesel,
                                                                                "num": num })[0].id);
    if ( relativeId.peselCheck      == yesNoOptions.yes.value ) {
        relativeId.pesel            = valueById(underscore.where(formArray, {   "prefix": textboxIdPrefix.relativeDocumentID,
                                                                                "num": num })[0].id);
        relativeId.idNum            = "";
        relativeId.idType           = "";
        relativeId.idId             = "";
    } else {
        relativeId.pesel            = "";
        relativeId.idNum            = valueById(underscore.where(formArray, {   "prefix": textboxIdPrefix.relativeDocumentID,
                                                                                "num": num })[0].id);
        relativeId.idType           = valueById(underscore.where(formArray, {   "prefix": textboxIdPrefix.relativeDocument2,
                                                                                "num": num })[0].id);
        relativeId.idId             = underscore.where(idOptions,           {   "value": relativeId.idType })[0].id;
    }
        relativeId.relationship     = valueById(underscore.where(formArray, {   "prefix": textboxIdPrefix.relativeRelationship,
                                                                                "num": num })[0].id);
        relativeId.relationshipId   = underscore.where(relationshipOptions, {   "name": relativeId.relationship })[0].code;
        relativeId.disability       = valueById(underscore.where(formArray, {   "prefix": textboxIdPrefix.relativeDisability,
                                                                                "num": num })[0].id);
        relativeId.disabilityId     = underscore.where(disabilityOptions,   {   "name": relativeId.disability })[0].code;
    if (relativeId.commonGround     == yesNoOptions.no.value) {
        relativeId.voivodeship      = valueById(underscore.where(formArray, {   "prefix": textboxIdPrefix.relativeVoivodeship,
                                                                                "num": num })[0].id);
        relativeId.community        = valueById(underscore.where(formArray, {   "prefix": textboxIdPrefix.relativeCommunity,
                                                                                "num": num })[0].id);
        relativeId.place            = valueById(underscore.where(formArray, {   "prefix": textboxIdPrefix.relativePlace,
                                                                                "num": num })[0].id);
        relativeId.postal           = valueById(underscore.where(formArray, {   "prefix": textboxIdPrefix.relativePostal,
                                                                                "num": num })[0].id);
        relativeId.street           = valueById(underscore.where(formArray, {   "prefix": textboxIdPrefix.relativeStreet,
                                                                                "num": num })[0].id);
        relativeId.flatNum          = valueById(underscore.where(formArray, {   "prefix": textboxIdPrefix.relativeFlatNumber,
                                                                                "num": num })[0].id);
        relativeId.apartmentNum     = valueById(underscore.where(formArray, {   "prefix": textboxIdPrefix.relativeApartmentNumber,
                                                                                "num": num })[0].id);
        relativeId.telephone        = valueById(underscore.where(formArray, {   "prefix": textboxIdPrefix.relativeTelephoneNumber,
                                                                                "num": num })[0].id);
    }
    return relativeId;
}

function valueById(id) {
    return document.getElementById(id).value;
}

function mapCreator(formArray) {
    var result = "";
    formArray.forEach(function (entry) {
        if (entry.type == "head" || entry.type == "head2") {

            result += "<a style=\"display: block;\" id=\"" 
                    + entry.formId 
                    + "Map" 
                    + "\" href='javascript:scrollToId(\"" 
                    + entry.formId + "\")'>" 
                    + entry.value 
                    + "<a/></br>";
        }
    });
    return result;
}

function scrollToId(id) {
    var divPosition = $('#' + id).offset();
    $('html, body').animate({ scrollTop: divPosition.top }, "slow");
}

function concatForm() {
    var generatedForm = [];
    var employeeIdentification, relativeIdentification, relativeAddress, location, generation;

    employeeIdentification = employeeIdentificationSection(1);
    employeeIdentification.forEach(function (entry) {
        generatedForm.push(entry);
    });

    var sectionsCount = parseInt(relativesCount) + 2;
    for (var i = 2; i < sectionsCount; i++) {
        relativeIdentification = relativeIdentificationSection(i);
        relativeIdentification.forEach(function (entry) {
            generatedForm.push(entry);
        });

        relativeAddress = relativeAddressSection(i);
        relativeAddress.forEach(function (entry) {
            generatedForm.push(entry);
        });
    }

    location = locationIdentificationSection(sectionsCount);
    location.forEach(function (entry) {
        generatedForm.push(entry);
    });

    generation = generationSection(sectionsCount + 1);
    generation.forEach(function (entry) {
        generatedForm.push(entry);
    });

    return generatedForm;
}

function rulesCreator(field) {
    var rulesString = "";
    rulesString += field.name + ": { ";
    rulesString += "minlength:" + field.minlength + ",";
    rulesString += "maxlength:" + field.maxlength + ",";

    if (field.required) {
        rulesString += "required:" + field.required + ",";
    }

    if (field.date) {
        rulesString += "date:" + field.date + ",";
    }

    if (field.digits) {
        rulesString += "digits:" + field.digits + ",";
        rulesString += "min:" + field.min + ",";
        rulesString += "max:" + field.max + ",";
    }

    rulesString += "}, ";

    return rulesString;
}

function rulesMaker(fields) {
    var rulesString = "";
    for (var i = 0; i < fields.length - 1; i++) {
        rulesString += rulesCreator(fields[i]);
    }
    return eval("({" + rulesString + "})");
}

function setContainersDisplayState(headState, relativesCountState, generatedContentState) {
    document.getElementById(containers.head).style.display = headState;
    document.getElementById(containers.relativesCount).style.display = relativesCountState;
    document.getElementById(containers.generatedContent).style.display = generatedContentState;
}

function getRelativesWorkbook(action) {
    return getWorkbookResponse().then(function (response) {
        xlsxOrgResponse = response;
        action();
    }).catch(function (err) {
        if (err == "Received a 404 HTTP code.") {
            alert("Nie można pobrać skoroszytu, spróbuj ponownie");
        } else {
            alert(err.message || err);
        }
        throw err;
    });
}

function getWorkbookResponse() {
    return new Promise(function (resolve, reject) {
        var req = new XMLHttpRequest();
        var url = urlInput;
        req.open("GET", url, true);
        req.responseType = "arraybuffer";
        req.onreadystatechange = function () {
            if (req.readyState === 4) {
                if (req.status === 200) {
                    resolve(req.response);
                } else {
                    reject("Received a " + req.status + " HTTP code.");
                }
            }
        };
        req.send();
    });
}

function generateEmptyBlob(empty, data) {
    return generate("blob", empty, data).then(function (blob) {
        var xlsxName = empty ? "PustyFormularz.xlsx" : "FormularzZDanymi.xlsx";
        if (window.navigator && window.navigator.msSaveOrOpenBlob) {
            window.navigator.msSaveOrOpenBlob(blob, xlsxName);
        } else {
            var url = window.URL.createObjectURL(blob);
            var a = document.createElement("a");
            document.body.appendChild(a);
            a.href = url;
            a.download = xlsxName;
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
        }
    }).catch(function (err) {
        alert(err.message || err);
        throw err;
    });
}

function getWorkbookFromResponse() {
    return new Promise(function (resolve, reject) {
        if (xlsxOrgResponse) {
            var xlsxResponse = new Uint8Array(xlsxOrgResponse);
            resolve(XlsxPopulate.fromDataAsync(xlsxResponse));
        } else {
            reject("Can't use xlsx response please refresh page.");
        }
    });
}

function generate(type, empty, data) {
    return getWorkbookFromResponse().then(function (workbook) {
        var workbook = workbook;
        if (!empty) {
            xlsxDataWrite(workbook, data);
        }
        return workbook.outputAsync(type);
    });
}

function xlsxDataWrite(workbook, data) {

    var rn = 0,
        sheet = workbook.sheet(0);

    rn += 2;
    sheet.cell(columns[6] + rn).value(data.employeeIdentification[0].relativesCount);

    rn += 2;
    writeContent(sheet, dateCleanSeparators(data.employeeIdentification[0].date), rn, 5);
    writeContent(sheet, data.employeeIdentification[0].id, rn, 29);

    rn += 4;
    writeContent(sheet, data.employeeIdentification[0].surname, rn, 5, 29);

    rn += 4;
    writeContent(sheet, data.employeeIdentification[0].name, rn, 5);

    if (data.relativeIdentification.length != 0) {
        for (var i = 0; i < data.relativeIdentification.length; i++) {
            rn += 4;
            writeContent(sheet, data.relativeIdentification[i].name, rn, 5);

            rn += 2;
            writeContent(sheet, data.relativeIdentification[i].surname, rn, 5);

            rn += 2;
            writeContent(sheet, data.relativeIdentification[i].pesel, rn, 5);
            writeContent(sheet, dateCleanSeparators(data.relativeIdentification[i].birth), rn, 17);
            writeContent(sheet, data.relativeIdentification[i].commonGround, rn, 26);

            rn += 2;
            writeContent(sheet, data.relativeIdentification[i].idNum, rn, 5);
            writeContent(sheet, data.relativeIdentification[i].idType, rn, 15);
            writeContent(sheet, data.relativeIdentification[i].idId, rn, 24);

            rn += 3;
            rn = writeChecker(sheet, data.relativeIdentification[i].relationshipId, relationshipOptions, 5, 6, rn, relationshipOptions.length);

            rn += 2;
            rn = writeChecker(sheet, data.relativeIdentification[i].disabilityId, disabilityOptions, 5, 6, rn, disabilityOptions.length);

            if (data.relativeIdentification[i].commonGround == yesNoOptions.no.value) {
                rn += 3;
                writeContent(sheet, data.relativeIdentification[i].voivodeship, rn, 5);

                rn += 2;
                writeContent(sheet, data.relativeIdentification[i].community, rn, 5);

                rn += 2;
                writeContent(sheet, data.relativeIdentification[i].place, rn, 5);
                writeContent(sheet, data.relativeIdentification[i].postal, rn, 28);

                rn += 2;
                writeContent(sheet, data.relativeIdentification[i].street, rn, 5);

                rn += 2;
                writeContent(sheet, data.relativeIdentification[i].flatNum, rn, 5);
                writeContent(sheet, data.relativeIdentification[i].apartmentNum, rn, 14);
                writeContent(sheet, data.relativeIdentification[i].telephone, rn, 20);
            } else { rn += 11; }
        }
    } else { rn += 4; }

    rn += 15;
    writeContent(sheet, data.locationIdentification[0].place, rn, 5, 22);

    rn += 2;
    writeContent(sheet, dateCleanSeparators(data.locationIdentification[0].date), rn, 5, 12);
}

function writeContent(sheet, text, row, point, br) {
    for (var i = 0; i < text.length; i++) {
        if (br && i == br) {
            row += 2;point = 5;
        }
        sheet.cell(columns[point++] + row).value(text.charAt(i).toUpperCase());
    }
}

function writeChecker(sheet, relationshipId, relationships, codeStart, checkerStart, row, checkersCount) {
    var checker = -1;
    for (var i = 0; i < relationships.length; i++) {
        if (relationshipId == relationships[i].code) {
            checker = i;
            break;
        } else {
            checker = -1;
        }
    }

    if (checker != -1) {
        row += checker;

        sheet.cell(columns[codeStart] + row).value(relationshipId);
        writeContent(sheet, "X", row, checkerStart);
        var toLast = 0;

        for (var i = 0; i < relationships.length; i++) {
            if (checker == i) {
                toLast = relationships.length - i;
                break;
            }
        }

        return row + toLast;
    } else {
        return checkersCount;
    }
}

function formPostGenerationPreparation(formArray) {

    var map = document.getElementById(containers.map);
    if (map) {
        map.innerHTML = mapCreator(formArray);
    }

    formArray.forEach(function (entry) {

        //Operations on from-group id
        if (entry.formId) {

            //documentOptionHide
            if (entry.formId.includes(formGroupIdPrefix.relativeDocument1) || entry.formId.includes(formGroupIdPrefix.relativeDocument2)) {
                document.getElementById(entry.formId).style.display = "none";
            }

            //peselActionAssign
            if (entry.formId.includes(formGroupIdPrefix.relativePesel)) {
                $("#" + entry.id).change(function () {

                    var num = entry.num;

                    if ($(this).val() == yesNoOptions.yes.value) {
                        document.getElementById(formGroupIdPrefix.relativeDocument1 + num).style.display = "block";
                        document.getElementById(formGroupIdPrefix.relativeDocument2 + num).style.display = "none";

                        $('#' + textboxIdPrefix.relativeDocumentID + num).rules('remove', 'maxlength');
                        $('#' + textboxIdPrefix.relativeDocumentID + num).rules('add', { maxlength: 11 });

                        document.querySelector("label[for=\"" + textboxIdPrefix.relativeDocumentID + num + "\"]").textContent = idOptions[0].name + " członka rodziny:";
                    } else {
                        document.getElementById(formGroupIdPrefix.relativeDocument1 + num).style.display = "none";
                        document.getElementById(formGroupIdPrefix.relativeDocument2 + num).style.display = "block";

                        $('#' + textboxIdPrefix.relativeDocumentID + num).rules('remove', 'maxlength');
                        $('#' + textboxIdPrefix.relativeDocumentID + num).rules('add', { maxlength: 9 });

                        document.querySelector("label[for=\"" + textboxIdPrefix.relativeDocumentID + num + "\"]").textContent = "Seria i nr " + idOptions[1].value + " / " + " seria i nr " + idOptions[2].value + "u" + " członka rodziny:";
                    }
                    if ($(this).val() == "") {
                        document.getElementById(formGroupIdPrefix.relativeDocument1 + num).style.display = "none";
                        document.getElementById(formGroupIdPrefix.relativeDocument2 + num).style.display = "none";

                        document.querySelector("label[for=\"" + textboxIdPrefix.relativeDocumentID + num + "\"]").textContent = "Pesel / seria i nr dow. os / seria i nr paszportu członka rodziny:";
                    }

                    document.getElementById(textboxIdPrefix.relativeDocumentID + num).value = "";
                });
            }

            //externalAddressOptionHide
            var externalAddressFormGroups = [formGroupIdPrefix.relativeAddressSectionTitle, formGroupIdPrefix.relativeVoivodeship, formGroupIdPrefix.relativeCommunity, formGroupIdPrefix.relativePlace, formGroupIdPrefix.relativePostal, formGroupIdPrefix.relativeStreet, formGroupIdPrefix.relativeFlatNumber, formGroupIdPrefix.relativeApartmentNumber, formGroupIdPrefix.relativeTelephoneNumber];
            for (var i = 0; i < externalAddressFormGroups.length; i++) {
                if (entry.formId.includes(externalAddressFormGroups[i])) {
                    document.getElementById(entry.formId).style.display = "none";
                }
            }

            //hideExternalAddressInMap
            if (map) {
                if (entry.formId.includes(externalAddressFormGroups[0])) {
                    document.getElementById(entry.formId + "Map").style.display = "none";
                }
            }

            //commonGroundActionAssign
            if (entry.formId.includes(formGroupIdPrefix.relativeCommonGround)) {
                $("#" + entry.id).change(function () {

                    var num = entry.num;
                    for (var i = 0; i < externalAddressFormGroups.length; i++) {
                        if ($(this).val() == yesNoOptions.no.value) {
                            document.getElementById(externalAddressFormGroups[i] + num).style.display = "block";
                            if (i == 0 && map) {
                                document.getElementById(externalAddressFormGroups[i] + num + "Map").style.display = "block";
                            }
                        } else {
                            document.getElementById(externalAddressFormGroups[i] + num).style.display = "none";
                            //hideInMapExternalAddress
                            if (i == 0 && map) {
                                document.getElementById(externalAddressFormGroups[i] + num + "Map").style.display = "none";
                            }
                        }

                        //clearExternalAddressValuesOnCommonGroundChange
                        if (externalAddressFormGroups[i].includes("TextboxFormGroupID")) {
                            var name = externalAddressFormGroups[i].split("TextboxFormGroupID")[0];
                            document.getElementById(name + "TextboxID" + num).value = "";
                        }
                    }
                });
            }
        }

        //emptySpreadsheetButtonAssign
        if (entry.id.includes("spreadsheetEmptyButton")) {
            $("#" + entry.id).click(function () {
                generateEmptyBlob(true);
            });
        }
    });
}

function dateToDDMMYYYY(date) {
    var d = date.getDate();
    var m = date.getMonth() + 1;
    var y = date.getFullYear();
    return '' + (d <= 9 ? '0' + d : d) + '-' + (m <= 9 ? '0' + m : m) + '-' + y;
}

function dateCleanSeparators(date) {
    return      date.charAt(0) + date.charAt(1) 
            +   date.charAt(3) + date.charAt(4) 
            +   date.charAt(6) + date.charAt(7) + date.charAt(8) + date.charAt(9);
}

$.validator.methods.date = function (value, element) {
    var bits = value.match(/([0-9]+)/gi),
        str;
    if (!bits) return this.optional(element) || false;
    str = bits[1] + '/' + bits[0] + '/' + bits[2];
    return this.optional(element) || !/Invalid|NaN/.test(new Date(str));
};

$.extend($.validator.messages, {
    required: "To pole jest wymagane.",
    remote: "Proszę o wypełnienie tego pola.",
    email: "Proszę o podanie prawidłowego adresu email.",
    url: "Proszę o podanie prawidłowego URL.",
    date: "Proszę o podanie prawidłowej daty.",
    dateISO: "Proszę o podanie prawidłowej daty (ISO).",
    number: "Proszę o podanie prawidłowej liczby.",
    digits: "Proszę o podanie samych cyfr.",
    creditcard: "Proszę o podanie prawidłowej karty kredytowej.",
    equalTo: "Proszę o podanie tej samej wartości ponownie.",
    extension: "Proszę o podanie wartości z prawidłowym rozszerzeniem.",
    maxlength: $.validator.format("Proszę o podanie nie więcej niż {0} znaków."),
    minlength: $.validator.format("Proszę o podanie przynajmniej {0} znaków."),
    rangelength: $.validator.format("Proszę o podanie wartości o długości od {0} do {1} znaków."),
    range: $.validator.format("Proszę o podanie wartości z przedziału od {0} do {1}."),
    max: $.validator.format("Proszę o podanie wartości mniejszej bądź równej {0}."),
    min: $.validator.format("Proszę o podanie wartości większej bądź równej {0}."),
    pattern: $.validator.format("Pole zawiera niedozwolone znaki.")
});

//Polyfills

if (!String.prototype.includes) {
  String.prototype.includes = function(search, start) {
    'use strict';
    if (typeof start !== 'number') {
      start = 0;
    }
    
    if (start + search.length > this.length) {
      return false;
    } else {
      return this.indexOf(search, start) !== -1;
    }
  };
}

Array.prototype.forEach = function (callback, thisArg) {
    if (typeof callback !== "function") {
        throw new TypeError(callback + " is not a function!");
    }
    var len = this.length;
    for (var i = 0; i < len; i++) {
        callback.call(thisArg, this[i], i, this);
    }
};