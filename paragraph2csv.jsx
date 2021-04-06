#target InDesign

var doc = app.activeDocument;
var sel = app.selection;
var allPages = app.activeDocument.pages;
var csvSep = "\t"; // разделителей полей в csv-файле
var htmlTaggedParagraph = ""; // создаем пустой параграф
var pageNum = app.activeDocument.pages[0].name,
      pageName = (pageNum * 1 >= 10) ? (pageNum) : ("0" + pageNum);
var cardCase = []; //Карточка дела

var edYear = "2020", // год издания
    edNumber = "50", // номер издания
    edID = 0, //  [0] - газета, [1] - приложение, [2] - приложение
    edName, csvCheck,
    dayMinusOne = 1, // вычитаем день из даты - 1 || нет - 0
    rubricName = sel[0].paragraphs[0].appliedParagraphStyle.name;


//Описываем HTML-теги
var htmlTag_Paragraph = ["<p>", "</p>"],
    htmlTag_Head3 = ["<h3>", "</h3>"],
    htmlTag_Bold = ["<b>", "</b>"],
    htmlTag_Italic = ["<i>", "</i>"],
    htmlTag_Div = ["<div>", "</div>"],
    htmlTag_BlockQuote = ["<blockquote>", "</blockquote>"],
    htmlTag_Sup = ["<sup>", "</sup>"],
    htmlTag_Sub = ["<sub>", "</sub>"],
    htmlTag_ListBullet = ["<ul>", "</ul>"],
    htmlTag_ListNumber = ["<ol>", "</ol>"],
    htmlTag_ListItem = ["<li>", "</li>"],
    htmlTag_AlignRight = ["<p style=\'text-align: right;\'>", htmlTag_Paragraph[1]],
    htmlTag_Line = "<hr>",
    htmlTag_TR = ["<tr>", "</tr>"],
    htmlTag_TD = ["<td>", "</td>"];


 var  twData, twEvent, twOrg, twWho, twPlace, twWhere, twTheme,
    htmlTag_TableWalk = "<table cellspacing=\"0\" cellpadding=\"0\" border=\"1\"><tbody><tr><td><p><b>"+twData+"</b></p></td>"+
        "<td><p class=\"bm_title\"><b>"+twEvent+"</b></p></td></tr><tr><td><p>"+twOrg+"</p></td><td><p>"+twWho+"</p></td>"+
        "</tr><tr><td><p>"+twPlace+"</p></td><td><p>"+twWhere+"</p></td></tr><tr><td><p><i>"+twTheme+"</i></p></td>"+
        "<td><br><p></p></td></tr></tbody></table>";
    
 var  tcDetails, tcDetailsName, tcFirstPerson, tcFirstPersonName, tcSecondPerson, tcSecondPersonName, tcRowsCount
    htmlTag_CardCase = "<table cellspacing=\"0\" cellpadding=\"0\" border=\"1\">"+
        "<tbody><tr><td rowspan= tcRowsCount >"+
        "<p><b>Карточка дела</b>"+
        "</p></td><td><p><b>"+ tcDetails +"</b>"+
        "</p></td><td><p>"+ tcDetailsName +"</p></td></tr>"+
        "<tr><td><p><b>"+ tcSecondPerson +"</b>"+
        "</p></td><td>"+ tcSecondPersonName +"</td></tr></tbody></table>";
    
//Параметры замены текста
var clearTextPar = {
    "\\t":" ",      //Табуляция на пробел
    "\\s":" ",      //Любой пробел на пробел
    "\\n":" ",       //Перевод строки на пробел
    "\\r":" ",      //Знак абзаца на пробел
    "\\u00a0":"",
    "­":"",         //Принудительный перенос
    "‑":"-",        //Недодефис меняем на дефис
    " +":" ",      //Несколько пробелов на один пробел
    "\\u00a":""
    };

var replaceTextPar = {
    "справка":"Справка",
    "пример":"Пример"
    }

var authorIDName = [
        ["8456" , "Аркадий", "Дмитриев"],
        ["56279", "Наталья", "Пешкова"]
];

//Ассоциативный массив
var fieldName = {
        "00-ID":"IE_XML_ID", // ID - без разницы какой, но должен отличаться в одном файле
        "01-Head":"IE_NAME", // Заголовок статьи
        "02-Active":"IE_ACTIVE", // Активна статья или нет "Y" || ''N" (N)
        "03-ActiveDate":"IE_ACTIVE_FROM", // Дата активации статьи (01.04.2020)
        "04-TextPreview":"IE_PREVIEW_TEXT", // Анонс статьи
        "05-TextPreviewType":"IE_PREVIEW_TEXT_TYPE", // Анонс Text/HTML (TEXT)
        "06-TextMain":"IE_DETAIL_TEXT", // Основной текст статьи
        "07-TextMainType":"IE_DETAIL_TEXT_TYPE", // Текст Text/HTML (HTML)
        "08-SortNum":"IE_SORT", // Порядко нумерации статьи (500)
        "09-MainTheme":"IP_PROP319", // "Руководителю", "Бухгалтеру", "Юристу", "Личное"
        "10-Author":"IP_PROP181", // ID Автора
        "11-Access":"IP_PROP114", // "(нет)", "открыт для всех", "для подписчиков", "всегда открыт"
        "12-PageNum":"IP_PROP271", // Номер полосы
        "13-Edition":"IC_GROUP0", // Выпуск
        "14-EditionYear":"IC_GROUP1", // Год выпуска / Раздел - уровень 1
        "15-EditionNum":"IC_GROUP2", // Номер выпуска / Раздел - уровень 2
        /* Свойства официальных документов*/
        "16-MainTheme":"IP_PROP359", // "Руководителю", "Бухгалтеру", "Юристу", "Личное" = prop319
        "17-DocNum":"IP_PROP339", // Номер документа
        "18-DateApproved":"IP_PROP338", // Дата принятия документа
        "19-DateApprovedMinust":"IP_PROP432", // Дата принятия документа Минюстом
        "20-DocNumMinust":"IP_PROP433", // Номер регистрации в Минюсте
        "21-DocType":"IP_PROP293", // Вид документа prop293
        "22-Access":"IP_PROP303", // "(нет)", "открыт для всех", "для подписчиков", "всегда открыт"
        /* Свойства Консультаций*/
        "23-MainTheme":"IP_PROP404", // "Руководителю", "Бухгалтеру", "Юристу", "Личное" = prop319
        "24-Author":"IP_PROP367", // ID Автора
        "25-Access":"IP_PROP369" // "(нет)", "открыт для всех", "для подписчиков", "всегда открыт"

};

//преобразуем ассоциативный массив в строку
var fieldNameString = AZ_arrayToString(fieldName, csvSep);

var prop319 = ["Руководителю", "Бухгалтеру", "Юристу", "Личное"],
      prop114 = ["(нет)", "открыт для всех", "для подписчиков", "всегда открыт"],
      propEdition = ["", "", ""],  
      prop293 = ["(не установлено)", "распоряжение", "постановление", "приказ", 
                      "письмо", "определение", "постановление Пленума", "постановление Президиума", 
                      "информационное письмо", "информационное письмо Президиума", "решение", 
                      "разъяснение", "вопрос-ответ", "проект решения", "проект постановления Пленума", 
                      "проект постановления ", "инструкция", "информационное сообщение", "положение", 
                      "постановление Правления", "проект приказа", "решение Комиссии", "указ", "указание", 
                      "методическое указание", "протокол"];
                  
var docDepartment = [
        ["Министерства финансов", "Министерство финансов РФ"], 
        ["Федеральной налоговой службы", "Федеральная налоговая служба"],
        ["ФСС", "Фонд социального страхования"],
        ["прав потребителей", "Федеральная служба по надзору в сфере защиты прав потребителей и благополучия человека"],
        ["экономического развития", "Министерство экономического развития РФ"],
        ["Центрального Банка", "Центральный Банк РФ"],
        ["ФСС", "Федеральная служба по труду и занятости РФ"],
        ["труда","Министерство труда и социальной защиты"]
];
      
// Создаем массив с пустыми полями
var csvText = new Array (AZ_arraySize(fieldName));
for (var i = 0; i < csvText.length; csvText[i++] = "");

csvText[0] = Math.floor(Math.random() * (200000)) + 800000; // Генерируем уникальный ID
csvText[2] = "Y"; // активный или нет
csvText[3] = AZ_getDateString(new Date(), ".")[1] + "  " + AZ_randomTime(); // дата активации
csvText[5] = "text"; // анонс
csvText[7] = "html"; // текст
csvText[8] = 500; // порядок сортировки
csvText[9] = prop319[edID]; // "Руководителю", "Бухгалтеру", "Юристу", "Личное"
csvText[11] = prop114[1]; // "(нет)", "открыт для всех", "для подписчиков", "всегда открыт"
csvText[12] = pageNum; // Номер полосы
csvText[13] = propEdition[edID]; 
csvText[14] = edYear; 
csvText[15] = edNumber;
csvText[16] = prop319[edID];
csvText[22] = prop114[1];
csvText[23] = prop319[edID];

//=========================================================
//Архив Статей
//=========================================================
function EG_Articles () {
    for (j=0; j<sel.length; j++) { //Перебираем выбранные фреймы
        if (sel[j] instanceof TextFrame) {
            csvText[12] = sel[j].parentPage.name; // Номер полосы для каждой статьи
        
            var selParagraphs = sel[j].parentStory.paragraphs;

            for (k = 0; k < selParagraphs.length; k++) { //Перебираем параграфы в выбранном фрейме

                switch (selParagraphs[k].appliedParagraphStyle.name) {

                     //Исключаем некоторые стили
                     case 'AT_рубрика':
                     case 'AT_отсыл':
                     case 'Рубрика':
                     case 'Начало':
                     case 'Окончание':
                     case 'Стр':
                     break;
                     
                    //Заголовки
                    case 'AZ_средний_заголовок':
                    case 'AZ_маленький_заголовок':
                    case 'AZ_заголовок':
                    case 'AZ_Заголовок':
                    case 'Колонка. Заголовок':
                    case 'Заголовок III':
                    case 'Заголовок II':
                    case 'Заголовок':
                    case 'AB_Бокс_вынос_заг':
                        
                           csvText[1] += AZ_clearText(selParagraphs[k], clearTextPar); 
                      
                    break;
                    
                    //Анонс
                    case 'AV_Врез_small':
                    case 'AV_Врез':
                    case 'Вопрос':
                    case 'Врез II':
                    case 'Вопрос':
                    case 'Врез':
                        csvText[4] = AZ_clearText(selParagraphs[k], clearTextPar);  
                    break;   
                    
                    //Подзаголовок 1
                    case 'Подзаголовок I':
                    case 'AT_подзаг':
                        htmlTaggedParagraph += htmlTag_Head3[0] + AZ_clearText(selParagraphs[k], clearTextPar) + htmlTag_Head3[1];
                    break;

                    //Подзаголовок 2
                    case 'Подзаголовок II':
                    case 'Подзаголовок':
                    case 'AB_Бокс_подзаг':
                        htmlTaggedParagraph += htmlTag_Paragraph[0] + htmlTag_Bold[0] + AZ_clearText( AZ_clearText( selParagraphs[k], clearTextPar), replaceTextPar) + htmlTag_Bold[1] + htmlTag_Paragraph[1];
                    break;                

                     //Сохраняем автора  
                     case 'AT_автор_текст':
                     case 'AT_автор_полосы':
                     case 'Автор на первой':
                     case 'Автор в тексте':
                     case 'Автор':
                     case 'ФИО, должность':
                        csvText[10] = AZ_authorID(selParagraphs[k].contents);
                     break;
                     
                    //Обрабатывам списки
                     case 'AT_текст_список':
                     case 'Текст. Список':
                     case 'Список':
                          htmlTaggedParagraph += htmlTag_ListBullet[0] + htmlTag_ListItem[0] + htmlTag_Paragraph[0] + AZ_clearText(selParagraphs[k], clearTextPar) + htmlTag_Paragraph[1] + htmlTag_ListItem[1] + htmlTag_ListBullet[1] ;  
                          htmlTaggedParagraph = htmlTaggedParagraph.replace ("</ul><ul>", ""); //Удаляем лишние теги
                     break;
                    
                    //Карточка дела
                    case 'Реквизиты':
                    case 'Истец/Ответчик':
                        cardCase.push(AZ_clearText(selParagraphs[k], clearTextPar).split("\t"));      
                     break;
                    
                    // Подписи
                    case 'Подпись':
                        htmlTaggedParagraph += htmlTag_AlignRight[0] + htmlTag_Italic[0] + AZ_clearText(selParagraphs[k], clearTextPar) + htmlTag_Italic[1] + htmlTag_AlignRight[1] + htmlTag_Line;
                    break;

                    case 'Текст. Подпись':
                        htmlTaggedParagraph += htmlTag_AlignRight[0] + htmlTag_Italic[0] + AZ_clearText(selParagraphs[k], clearTextPar) + htmlTag_Italic[1] + htmlTag_AlignRight[1];
                    break;      
                    
                    default:
                          htmlTaggedParagraph += htmlTag_Paragraph[0] + AZ_clearText(selParagraphs[k], clearTextPar) + htmlTag_Paragraph[1]; 

                    break;
                }// End SWITCH
                //$.writeln(htmlTaggedParagraph)
            }//End FOR selParagraphs
        }//End IF constructorname 
    }//End FOR sel
    //Сохраняем текст и добавляем анонс к основному тексту
    csvText[6] = htmlTag_Paragraph[0] + csvText[4] + htmlTag_Paragraph[1] + /*cardCase.join("#") +*/ htmlTaggedParagraph;
    
    // Пишем массив csvText в файл
    csv.open ("a");
    alert(csvText.join(csvSep));
    csv.write(csvText.join(csvSep));
    csv.close();

    alert(csvPath);
}// end main




//=================================================================
// ЗАПУСК всего!!!!!!!!!!!!!!!!!!!!!!!!!!
//=================================================================
var csvPath = Folder.desktop + pageName + "_" + AZ_getDateString(new Date(), "-")[0] + "_" + ".txt";
var csv = File (csvPath);
        csv.open ("a");
        csv.writeln (fieldNameString); // Записываем первую строку в csv
        csv.close();


    EG_Articles()

//============================================
// Функция перевода строки в дату
//============================================
function AZ_getDateString(dateString, sep) {
	var date = dateString || new Date();
	var day = (date.getDate() > 10 ) ? date.getDate() : ("0" + date.getDate() ),
		month = ((date.getMonth()+1) >= 10 ) ? (date.getMonth()+1) : ("0" + (date.getMonth()+1) ),
		year = (date.getFullYear() > 10 ) ? date.getFullYear() : ("0" + date.getFullYear() ),
		hours = (date.getHours() > 10 ) ? date.getHours() : ("0" + date.getHours() ),
		minutes = (date.getMinutes() > 10 ) ? date.getMinutes() : ("0" + date.getMinutes() );
	day = day-dayMinusOne;
	return arr = [year + sep + month + sep + day, day + sep + month + sep + year];
};

//============================================
//Функция преобразования ассоциативного массива в строку 
//============================================
function AZ_arrayToString(obj, sep) {
    var arr = [], p, i = 0;
    for (p in obj)
        arr.push(obj[p]);
    return arr.join(sep);
}

//============================================
//Функция определения размера ассоциативного массива 
//============================================
function AZ_arraySize(obj) {
    var size = 0, key;
    for (key in obj) {
        if (obj.hasOwnProperty(key)) size++;
    }
    return size;
};

//============================================
// Чистим текст от всякой ерунды
// Принимаем в себя текст и Ассоциативный массив с параметрами замен
//============================================
function AZ_clearText (txt, array) {
    var text, textReady;
    (txt instanceof Object) ? (text = txt.contents) : (text = txt) //проверяем что получаем на входе
    
    for (key in array) {
         textReady = text.replace (new RegExp( key, "g" ), array[key]);
         text = textReady;
    } 
    return textReady;
}

//============================================
// Получаме номер документа из строки 
//============================================
function AZ_docNum (docNum) {
    var text = docNum.match(/№.+/);
    return text[0].replace(/№\s/, "");
 }

//============================================
// Получаме дату документа вида "01.01.2020" из "1 января 2020" 
//============================================
function AZ_docDate (docDate) {
    var month = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря"];
    var text = docDate.match(/от\s\d+.+?\d{4}/);
    text[0] = text[0].replace(/от\s/, "");
    var date = text[0].split(" ");
    (date[0] < 10) ? (date[0] = "0" + date[0]) : (false);
    for (i=0; i<month.length; i++) {
        if (month[i] == date[1]) {
            (i >=9) ? (date[1] = i + 1) : (date[1] = "0" + (i + 1));
            return date.join(".");
        }
     }
 }


//============================================
// Получаме ID автора из массива
//============================================
function AZ_authorID (author) {
    for (i = 0;  i < authorIDName.length; i++) {
        if (author.match(new RegExp (authorIDName[i][2], "i")) ) {
            return authorIDName[i][0];
        }
    } return "";
}

//============================================
// Получаме вид документа из массива
//============================================
function AZ_docType (docType) {
    for (i = 0;  i < prop293.length; i++) {
        if (docType.match(new RegExp (prop293[i], "i")) ) {
            return prop293[i]
        }
    }
}

//============================================
// Получаме вид документа из массива
//============================================
function AZ_docDep (department) {
    for (i = 0;  i < docDepartment.length; i++) {
        if (department.match(new RegExp (docDepartment[i][0], "i")) ) {
            return docDepartment[i][1];
        }
    }
}


//============================================
// Генерируем рандомно время от 13 до 19
//============================================
function AZ_randomTime() {
        var curDate = new Date();
        var sec, minute, hour;
        var curHour = curDate.getHours();
        var curMinutes = curDate.getMinutes();
        sec = Math.floor(10 + Math.random() * (59 - 10));
        minute = Math.floor(Math.random() * (curMinutes));
        hour = Math.floor(10 + Math.random() * (curHour +1 - 10));
        //hour += 10; 
        (minute <10) ? (minute = "0" + minute) : (false);
        (hour <10) ? (hour = "0" + hour) : (false);
        return hour + ":" + minute + ":" + sec;
}
