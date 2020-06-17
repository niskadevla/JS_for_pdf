// PDF with different page sizes processing v 1.1
// Sort pages based on page size and makes report with info about sizes
// Report page sizes in current PDF
// Split pages By size
// How to use: // put this file to your ..\Acrobat\Javascripts folder and restart Acrobat.
//Click tools
// You'll see 5 new menu items in Tool menu.
// Copyright (c) Bielishev Denis
// niskadevla@gmail.com

/**
 * @param {string} CURRENCY                      currency (UAH)
 * @CONFIG {object} CONFIG    object of prices   (UAH)
 *    @param {number} CONFIG.FOLD                folding price (UAH)
 *    @param {number} CONFIG.CUT                 cuting price (UAH)
 *    @param {number} CONFIG.SORT.RANGE          the price of one range (UAH)
 *    @param {number} CONFIG.SORT.FORMAT         the price of one format (UAH)
*/

const CURRENCY = "UAH"; // валюта грн.

// config cost (uah)
const CONFIG = {
  FOLD: 4,
  CUT: 4,
  SORT: {
    RANGE: 4,
    FORMAT:4
  }
};

// const for fold of price
const STAMP_A4 = 192, // min folding width // Штамп-размер (верхнего слоя сложенного чертежа) - сложение под А4
      STAMP_A3 = 402,  // min folding width // Штамп-размер (верхнего слоя сложенного чертежа) - сложение под А3
      F = 20, // binding field // поле-корешок
      KH = 300; // fold height // коэффициент по высоте (макс. расстояние до следующего сгиба)


if (app.viewerVersion < 10) {
    app.addMenuItem({
        cName: "Pages sorter",
        cUser: "Pages sorter",
        cParent: "Tools",
        cExec: "PagesSorter()",
        cEnable: "event.rc=(event.target != null);"
    });
    app.addMenuItem({
        cName: "Pages size",
        cUser: "Pages size",
        cParent: "Tools",
        cExec: "PagesSize()",
        cEnable: "event.rc=(event.target != null);"
    });
    app.addMenuItem({
        cName: "Split by Pages size",
        cUser: "Split by Pages size",
        cParent: "Tools",
        cExec: "SplitByPagesSize()",
        cEnable: "event.rc=(event.target != null);"
    });
    app.addMenuItem({
        cName: "Price of print and folding, cuting to A4",
        cUser: "Price of print and folding, cuting to A4",
        cParent: "Tools",
        cExec: "PricePrintFoldCutToA4()",
        cEnable: "event.rc=(event.target != null);"
    });
    app.addMenuItem({
        cName: "Price of print and folding, cuting to A3",
        cUser: "Price of print and folding, cuting to A3",
        cParent: "Tools",
        cExec: "PricePrintFoldCutToA3()",
        cEnable: "event.rc=(event.target != null);"
    });

} else {
    app.addToolButton({
        cName: "Pages sorter",
        cLabel: "Pages sorter",
        cExec: "PagesSorter()",
        cEnable: "event.rc=(event.target != null);"
    });
    app.addToolButton({
        cName: "Pages size",
        cLabel: "Pages size",
        cExec: "PagesSize()",
        cEnable: "event.rc=(event.target != null);"
    });
    app.addToolButton({
        cName: "Split by Pages size",
        cLabel: "Split by Pages size",
        cExec: "SplitByPagesSize()",
        cEnable: "event.rc=(event.target != null);"
    });
    app.addToolButton({
        cName: "Price of print and folding, cuting to A4",
        cLabel: "Price of print and folding, cuting to A4",
        cExec: "PricePrintFoldCutToA4()",
        cEnable: "event.rc=(event.target != null);"
    });
    app.addToolButton({
        cName: "Price of print and folding, cuting to A3",
        cLabel: "Price of print and folding, cuting to A3",
        cExec: "PricePrintFoldCutToA3()",
        cEnable: "event.rc=(event.target != null);"
    });
}

function createRepHeader(doc, rep2Head) {
    rep2Head.writeText("// Pages from " + doc.path + " sorted by sizes \n// All sizes in mm, based on Crop box\n");
}
PagesSorter = app.trustedFunction(function() {
    app.beginPriv();
    var aa = [];
    var acp = [];
    var cp, tmp, cc;
    var nr = 2.834;
    var prec = 0;
    var linePerpage = 59;
    var ind = 25;
    if (this.numPages < 2) {
        app.alert("Must be 2 or more pages in file");
        return -1;
    }
    var npw = 0;
    var nph = 0;
    var rep = new Report();
    rep.size = 1.2;
    rep.writeText("PDF page by size sorter (C) Denis 2019 niskadevla@gmail.com");
    createRepHeader(this, rep);
    for (cc = 0; cc < this.numPages; cc++) {
        aa.length = 0;
        for (cp = cc; cp < this.numPages; cp++) {
            ar = this.getPageBox("Crop", cp);
            widthP = (ar[2] - ar[0]) / nr;
            heightP = (ar[1] - ar[3]) / nr;
            if (heightP > widthP) {
                tmp = heightP;
                heightP = widthP;
                widthP = tmp;
            }
            aa.push({
                page: cp,
                width: widthP.toFixed(prec),
                height: heightP.toFixed(prec)
            });
        };
        aa.sort(function(a, b) {
            if (a.width == b.width) {
                return a.height - b.height;
            } else {
                return a.width - b.width;
            }
        });
        this.movePage(aa[0].page, cc - 1);
        if ((aa[0].width != npw) || (aa[0].height != nph)) {
            npw = aa[0].width;
            nph = aa[0].height;
            acp.push({
                page: cc,
                width: aa[0].width,
                height: aa[0].height
            });
        }
    }

    rep.writeText("// info about sizes for printing");
    rep.writeText("// start page end page width (mm) height (mm)");
    rep.indent(ind);
    for (tmp = 0; tmp < acp.length - 1; tmp++) {
        rep.writeText((acp[tmp].page + 1) + "\t\t" + acp[tmp + 1].page + "\t\t" + acp[tmp].width + "\t\t" + acp[tmp].height);
        if ((cc % linePerpage == 0) && (cc != 0)) {
            rep.breakPage();
            createRepHeader(this, rep);
        }
    }
    rep.writeText((acp[acp.length - 1].page + 1) + "\t\t" + this.numPages + "\t\t" + acp[tmp].width + "\t\t" + acp[tmp].height);
    rep.outdent(ind);
    rep.writeText("// ATTENTION: PAGE ORDER in file WAS CHANGED. Be careful on saving or save file with another name!");
    rep.writeText("// This script treat all pages as landscapes but NOT rotate pages in file. \n// Don't forget to turn on checkbox 'Rotate and center image' on Print dialog or rotate pages by Yourself");
    rep.writeText("// This scipt works 'as is' and no any warranties on working");
    rep.writeText("// End of report");
    var repPath = this.path.replace("Pages sizes report");
    var reportDoc = rep.open(repPath);
    reportDoc.info.Title = "Sorted pages by sizes in " + this.path;
    reportDoc.info.Producer = "Denis niskadevla@gmail.com" + this.path;
    app.endPriv();
    return;
});

function createRepSizeHeader(doc, rep2Head) {
    rep2Head.writeText("// Sizes of each page in " + doc.path + ". \n// All sizes in mm, based on Crop box\n// pagenum\twitdh\theight");
};
PagesSize = app.trustedFunction(function() {
    app.beginPriv();
    var cp;
    var nr = 2.834;
    var prec = 0;
    var linePerpage = 59;
    var ind = 25;
    var rep = new Report();
    rep.size = 1.2;
    rep.writeText("PDF pages size (C) Denis 2019 niskadevla@gmail.com");
    createRepSizeHeader(this, rep);
    rep.indent(ind);
    var count = 0;
    var prevWidthP = 0;
    var prevHeightP = 0;

    for (cp = 0; cp < this.numPages; cp++) {
        ar = this.getPageBox("Crop", cp);
        widthP = (ar[2] - ar[0]) / nr;
        heightP = (ar[1] - ar[3]) / nr;
        rep.writeText((cp + 1) + "\t" + widthP.toFixed(prec) + "\t" + heightP.toFixed(prec));

        if ((widthP != prevWidthP) || (prevHeightP != heightP)) {count++;};
        prevWidthP = widthP;
        prevHeightP = heightP;

        if ((cp % linePerpage == 0) && (cp != 0)) {
            rep.breakPage();
            createRepSizeHeader(this, rep);
        }
    }
    rep.writeText ("// End of report. Amount range: " + count);
    rep.outdent(ind);
    var repPath = this.path.replace("Pages sizes report");
    var reportDoc = rep.open(repPath);
    reportDoc.info.Title = "Sizes of each page in " + this.path;
    reportDoc.info.Producer = "Denis niskadevla@gmail.com" + this.path;
    app.endPriv();
    return;
});

function createRepSplitHeader(doc, rep2Head) {
    rep2Head.writeText("// Splitted files stored in source file folder" + ". \n// All sizes in mm, based on Crop box\n// cfn\t\twitdh\t\theight\t\tpage count");
}
SplitByPagesSize = app.trustedFunction(function() {
    app.beginPriv();
    var aa = [];
    var acp = [];
    var cp, tmp, cc;
    var nr = 2.834;
    var prec = 0;
    var linePerpage = 59;
    var ind = 25;
    if (this.numPages < 2) {
        app.alert("Must be 2 or more pages in file");
        return -1;
    }
    var npw = 0;
    var nph = 0;
    var rep = new Report();
    rep.size = 1.2;
    rep.writeText("PDF split by page size (C) Denis 2019 niskadevla@gmail.com");
    createRepSplitHeader(this, rep);
    for (cc = 0; cc < this.numPages; cc++) {
        aa.length = 0;
        for (cp = cc; cp < this.numPages; cp++) {
            ar = this.getPageBox("Crop", cp);
            widthP = (ar[2] - ar[0]) / nr;
            heightP = (ar[1] - ar[3]) / nr;
            if (heightP > widthP) {
                tmp = heightP;
                heightP = widthP;
                widthP = tmp;
            }
            aa.push({
                page: cp,
                width: widthP.toFixed(prec),
                height: heightP.toFixed(prec)
            });
        }
        aa.sort(function(a, b) {
            if (a.width == b.width) {
                return a.height - b.height;
            } else {
                return a.width - b.width;
            }
        });
        this.movePage(aa[0].page, cc - 1);
        if ((aa[0].width != npw) || (aa[0].height != nph)) {
            npw = aa[0].width;
            nph = aa[0].height;
            acp.push({
                page: cc,
                width: aa[0].width,
                height: aa[0].height
            });
        }
    }
    //console.println("acp.length = " + acp.length);
    rep.indent(ind);
    var cfn = this.path.replace(".pdf", "");
    var cen;
    for (tmp = 0; tmp < acp.length - 1; tmp++) {
        if ((cc % linePerpage == 0) && (cc != 0)) {
            rep.breakPage();
            createRepSplitHeader(this, rep);
        }
        try {
            cen = cfn + "_" + acp[tmp].width + "x" + acp[tmp].height + "_" + (acp[tmp + 1].page - acp[tmp].page) + ".pdf";
            this.extractPages({
                nStart: acp[tmp].page,
                nEnd: acp[tmp + 1].page - 1,
                cPath: cen
            });
        } catch (e) {
            console.println("Aborted: " + e)
        }
        rep.writeText(cen + "\t\t" + acp[tmp].width + "\t\t" + acp[tmp].height + "\t\t" + (acp[tmp + 1].page - acp[tmp].page));
    }
    try {
        this.extractPages({
            nStart: acp[acp.length - 1].page,
            nEnd: this.numPages - 1,
            cPath: cfn + "_" + acp[acp.length - 1].width + "x" + acp[acp.length - 1].height + "_" + (this.numPages - acp[acp.length - 1].page) + ".pdf"
        });
    } catch (e) {
        console.println("Aborted: " + e)
    }
    rep.writeText(cen + "\t\t" + acp[acp.length - 1].width + "\t\t" + acp[acp.length - 1].height + "\t\t" + (this.numPages - acp[acp.length - 1].page));
    rep.outdent(ind);
    rep.writeText("// ATTENTION: PAGE ORDER in source file WAS CHANGED. Be careful on saving or save file with another name!");
    rep.writeText("// This script treat all pages as landscapes but NOT rotate pages in file. \n// Don't forget to turn on checkbox 'Rotate and center image' on Print dialog or rotate pages by Yourself");
    rep.writeText("// This scipt works 'as is' and no any warranties on working");
    rep.writeText("// End of report");
    var repPath = this.path.replace("Pages sizes split report");
    var reportDoc = rep.open(repPath);
    reportDoc.info.Title = "Spit pages by sizes in " + this.path;
    reportDoc.info.Producer = "Denis niskadevla@gmail.com" + this.path;
    app.endPriv();
    return;
});



function createRepPriceHeader(doc, rep2Head) {
    rep2Head.writeText("// Sizes of each page in " + doc.path + ". \n// All sizes in mm, based on Crop box\n// pagenum\twitdh\theight");
};
PricePrintFoldCutToA4 = app.trustedFunction(function() {
    app.beginPriv();
    var aa = [];
    var abf = []; // array big formats
    var acp = [];
    var ar = [];
    var cp, tmp, cc;

    var KW = STAMP_A4 * 2 + 1; // коэффициент по горизонтали (макс. расстояние до следующего сгиба)

    // s - расстояние между штампом и корешком
    // sW - количество отрезков, максимальный размер (ширина) которых не превышает двух размеров
    // верхнего слоя
    // foldW - количество сгибов по горизонтали
    // sH - количество отрезков, максимальный размер (высота) которых не превышает двух размеров
    // верхнего слоя
    // foldH - количество сгибов по вертикали
    // fold - общее количество сгибов
    // cut - общее количество подрезов
    var s, sW, foldW, sH, foldH,
        fold = cut = 0;

    if (this.numPages < 2) {
        app.alert("Must be 2 or more pages in file");
        return -1;
    }

    var npw = 0;
    var nph = 0;
    var nr = 2.834;
    var prec = 0;
    var linePerpage = 59;
    var ind = 25;
    var rep = new Report();
    rep.size = 1.2;
    rep.color = color.black;

        // rep.writeText("PDF pages size (C) Denis 2019 niskadevla@gmail.com");
    createRepPriceHeader(this, rep);
    rep.indent(ind);
    var range = 0;
    var prevWidthP = 0;
    var prevHeightP = 0;

    // * Выводит все форматы на экран
    // Cчитает количество дипазонов
    for (cp = 0; cp < this.numPages; cp++) {
        ar = this.getPageBox("Crop", cp);
        widthP = (ar[2] - ar[0]) / nr;
        heightP = (ar[1] - ar[3]) / nr;

        if (heightP > widthP) {
            tmp = heightP;
            heightP = widthP;
            widthP = tmp;
        }
        rep.writeText((cp + 1) + "\t" + widthP.toFixed(prec) + "\t" + heightP.toFixed(prec));

            // Подсчет количества диапазонов
        if ((widthP != prevWidthP) || (heightP != prevHeightP)) {
          range++;
        }
        prevWidthP = widthP;
        prevHeightP = heightP;

        if ((cp % linePerpage == 0) && (cp != 0)) {
            rep.breakPage();
            createRepPriceHeader(this, rep);
        }
    } // **


    // Фильтрация листов, т.е. отбрасывание А4
    // Считаем количество форматов
    aa.length = 0;
    for (cc = 0; cc < this.numPages; cc++) {
      ar = this.getPageBox("Crop", cc);
      widthP = (ar[2] - ar[0]) / nr;
      heightP = (ar[1] - ar[3]) / nr;

          // формируем массив больших форматов, т.е. больше чем А4
      if ( (widthP > 211) || (heightP > 298) ) {
        abf.push({
          width: +widthP.toFixed(prec),
          height: +heightP.toFixed(prec)
        });
      }

          // Делаем все листы landscape, т.е. если нужно переварачиваем (виртуально)
      if (heightP > widthP) {
          tmp = heightP;
          heightP = widthP;
          widthP = tmp;
      }

          // округляем ширину и высоту
          // toFixed - округляет и превращает в строку
      widthP = widthP.toFixed(prec);
      heightP = heightP.toFixed(prec);

          // формируем массив разных форматов, т.е его длина и будет количеством форматов в документе
          // Заносим в массив первый размер, т.е. первую страницу
      if ( !aa.length ) {
        aa.push({
          width: widthP,
          height: heightP
        });
      }
          // Заносим в массив остальные размеры
      var counter = 0;
      for (var i = 0; i < aa.length; i++) {
        if (aa[i].width == widthP && aa[i].height == heightP) break;
        if (aa[i].width != widthP || aa[i].height != heightP) {
          counter++;
        }
        if ( (aa[i].width != widthP || aa[i].height != heightP) &&  counter == aa.length ) {
          aa.push({
            width: widthP,
            height: heightP
          });
        }
      }
    } // **

    var format = aa.length;


    // * Считаем сгибы
    console.println(abf.length);
    for (var i = 0; i < abf.length; i++) {
      s = abf[i].width - F - STAMP_A4;
      sW = Math.ceil(s / KW);
      foldW = sW * 2;
      sH = Math.floor(abf[i].height / KH);
      foldH = sH + 1;

      if (!sH) foldH = 0;

      fold += foldW + foldH;

      // Если больше чем А2, то cut = 5 (подрезов);
      // Если больше чем А3, то cut = 4 (подреза);
      cut += (abf[i].height > 421 && abf[i].width > 421) ? 5 : ( abf[i].height > 300 && abf[i].width > 300 ) ? 4 : 0;
    }

    // Общий подсчет
    var totalSum = 0;
    totalSum = fold * CONFIG.FOLD + cut * CONFIG.CUT + range * CONFIG.SORT.RANGE + format * CONFIG.SORT.FORMAT;

    rep.writeText ("// To A4 " + ". \n" +
                    " Amount ranges: " + range + ". \n" +
                    "Amount formats: " + format + ". \n" +
                    "Amount folds: " + fold + ". \n" +
                    "Amount cutings: " + cut + ". \n" +
                    "Total Sum: " + totalSum + " " + CURRENCY);
    rep.outdent(ind);
    var repPath = this.path.replace("Pages sizes report");
    var reportDoc = rep.open(repPath);
    reportDoc.info.Title = "Sizes of each page in " + this.path;
    reportDoc.info.Producer = "Denis niskadevla@gmail.com" + this.path;
    app.endPriv();
    return;
});



function createRepPriceHeader(doc, rep2Head) {
    rep2Head.writeText("// Sizes of each page in " + doc.path + ". \n// All sizes in mm, based on Crop box\n// pagenum\twitdh\theight");
};
PricePrintFoldCutToA3 = app.trustedFunction(function() {
    app.beginPriv();
    var aa = [];
    var abf = []; // array big formats
    var acp = [];
    var ar = [];
    var cp, tmp, cc;

    var KW = STAMP_A3 * 2 + 1; // коэффициент по горизонтали (макс. расстояние до следующего сгиба)
    var s, sW, foldW, sH, foldH,
        fold = cut = 0;

    if (this.numPages < 2) {
        app.alert("Must be 2 or more pages in file");
        return -1;
    }

    var npw = 0;
    var nph = 0;
    var nr = 2.834;
    var prec = 0;
    var linePerpage = 59;
    var ind = 25;
    var rep = new Report();
    rep.size = 1.2;
    rep.color = color.black;

        // rep.writeText("PDF pages size (C) Denis 2019 niskadevla@gmail.com");
    createRepPriceHeader(this, rep);
    rep.indent(ind);
    var range = 0;
    var prevWidthP = 0;
    var prevHeightP = 0;

    // * Выводит все форматы на экран
    // Cчитает количество дипазонов
    for (cp = 0; cp < this.numPages; cp++) {
        ar = this.getPageBox("Crop", cp);
        widthP = (ar[2] - ar[0]) / nr;
        heightP = (ar[1] - ar[3]) / nr;

        if (heightP > widthP) {
            tmp = heightP;
            heightP = widthP;
            widthP = tmp;
        }
        rep.writeText((cp + 1) + "\t" + widthP.toFixed(prec) + "\t" + heightP.toFixed(prec));

            // Подсчет количества диапазонов
        if ((widthP != prevWidthP) || (heightP != prevHeightP)) {
          range++;
        }
        prevWidthP = widthP;
        prevHeightP = heightP;

        if ((cp % linePerpage == 0) && (cp != 0)) {
            rep.breakPage();
            createRepPriceHeader(this, rep);
        }
    } // **


    // Фильтрация листов, т.е. отбрасывание А4
    // Считаем количество форматов
    aa.length = 0;
    for (cc = 0; cc < this.numPages; cc++) {
      ar = this.getPageBox("Crop", cc);
      widthP = (ar[2] - ar[0]) / nr;
      heightP = (ar[1] - ar[3]) / nr;

          // формируем массив больших форматов, т.е. больше чем А3
      if ( (widthP > 421) || (heightP > 298) ) {
        abf.push({
          width: +widthP.toFixed(prec),
          height: +heightP.toFixed(prec)
        });
      }

          // Делаем все листы landscape, т.е. если нужно переварачиваем (виртуально)
      if (heightP > widthP) {
          tmp = heightP;
          heightP = widthP;
          widthP = tmp;
      }

          // округляем ширину и высоту
          // toFixed - округляет и превращает в строку
      widthP = widthP.toFixed(prec);
      heightP = heightP.toFixed(prec);

          // формируем массив разных форматов, т.е его длина и будет количеством форматов в документе
          // Заносим в массив первый размер, т.е. первую страницу
      if ( !aa.length ) {
        aa.push({
          width: widthP,
          height: heightP
        });
      }
          // Заносим в массив остальные размеры
      var counter = 0;
      for (var i = 0; i < aa.length; i++) {
        if (aa[i].width == widthP && aa[i].height == heightP) break;
        if (aa[i].width != widthP || aa[i].height != heightP) {
          counter++;
        }
        if ( (aa[i].width != widthP || aa[i].height != heightP) &&  counter == aa.length ) {
          aa.push({
            width: widthP,
            height: heightP
          });
        }
      }
    } // **

    var format = aa.length;


    // * Считаем сгибы
    console.println(abf.length);
    for (var i = 0; i < abf.length; i++) {
      s = abf[i].width - F - STAMP_A3;
      sW = Math.ceil(s / KW);
      foldW = sW * 2;
      sH = Math.floor(abf[i].height / KH);
      foldH = sH + 1;

      if (!sH) foldH = 0;

      fold += foldW + foldH;

      // Если больше чем А2, то cut = 5 (подрезов);
      // Если больше чем А3, то cut = 4 (подреза);
      cut += ( abf[i].height > 421 && abf[i].width > 421 ) ? 5 : ( abf[i].height > 300 && abf[i].width > 300 ) ? 4 : 0;
    }

    // Общий подсчет
    var totalSum = 0;
    totalSum = fold * CONFIG.FOLD + cut * CONFIG.CUT + range * CONFIG.SORT.RANGE + format * CONFIG.SORT.FORMAT;

    rep.writeText ("// To A3" + ". \n" +
                    "Amount ranges: " + range + ". \n" +
                    "Amount formats: " + format + ". \n" +
                    "Amount folds: " + fold + ". \n" +
                    "Amount cutings: " + cut + ". \n" +
                    "Total Sum: " + totalSum + " " + CURRENCY);
    rep.outdent(ind);
    var repPath = this.path.replace("Pages sizes report");
    var reportDoc = rep.open(repPath);
    reportDoc.info.Title = "Sizes of each page in " + this.path;
    reportDoc.info.Producer = "Denis niskadevla@gmail.com" + this.path;
    app.endPriv();
    return;
});
