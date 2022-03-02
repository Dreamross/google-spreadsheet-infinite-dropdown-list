function SmartDataValidation(event) {
  //--------------------------------------------------------------------------------------
  // The event handler, adds data validation for the input parameters
  //--------------------------------------------------------------------------------------
  // Изменить настройки:
  //--------------------------------------------------------------------------------------

  var TargetSheet = {
    1: [
      {
        listName: 'Дополнительный', // name of the sheet where you want to perform data validation
        LogSheet: 'Данные', // sheet name, with data for dropdown lists
        NumOfLevels: 4, // number of dropdown list levels
        lcol: 1, // column number on the left, from which the first list begins; A=1, B=2, etc.
        lrow: 2, // line number, starting from which the list is triggered
        dataCol: 1, // column from which to take data
        KudaCol: 6, // column in which data is written to be replaced
        kudaRow: 2, // line from which the data to be replaced starts to be written
      },
    ],
    // If you need more:
    // Если нужно еще:
    2: [
      {
        listName: 'Дополнительный', // имя листа, где нужно осуществлять проверку данных
        LogSheet: 'Данные', // имя листа, с данными для выпадающиих списков
        NumOfLevels: 4, // количество уровней выпадающего списка
        lcol: 6, // номер колонки слева, с которой начинается первый список; A = 1, B = 2, etc.
        lrow: 2, // номер строки, начиная с которой срабатывает список
        dataCol: 13, // колонка из которой брать данные
        KudaCol: 18, // колонка в которую записываются данные для замены
        kudaRow: 2, // строка с которой начинают записываться данные для замены
      },
    ],
    3: [
      {
        listName: 'Основной',
        LogSheet: 'Данные',
        NumOfLevels: 4,
        lcol: 6,
        lrow: 2,
        dataCol: 13,
        KudaCol: 18,
        kudaRow: 2,
      },
    ],
    4: [
      {
        listName: 'Основной',
        LogSheet: 'Данные',
        NumOfLevels: 4,
        lcol: 1,
        lrow: 2,
        dataCol: 1,
        KudaCol: 6,
        kudaRow: 2,
      },
    ],
  };

  // =====================================================================================

  var FormulaSplitter = ';'; // depends on regional setting, ';' or ',' works for US
  //--------------------------------------------------------------------------------------

  //	====================================  key variables	 =================================

  // [ 01 ].Track sheet on which an event occurs

  var eventSource = event.source;
  var ts = eventSource.getActiveSheet();
  var sname = ts.getName();
  var activeCell = eventSource.getActiveRange().getColumn();

  let listName = '';
  let variant = '';
  var keys = Object.keys(TargetSheet);

  for (let z = 0; z <= keys.length; z += 1) {
    let value = keys[z];
    let tmpName = TargetSheet[value][0].listName;
    let colInit = TargetSheet[value][0].lcol;
    let colEnd = colInit + TargetSheet[value][0].NumOfLevels;

    if (sname === tmpName && colInit <= activeCell && activeCell < colEnd) {
      listName = tmpName;
      variant = value;
      break;
    }
  }

  if (sname === listName) {
    var { LogSheet } = TargetSheet[variant][0];
    var { NumOfLevels } = TargetSheet[variant][0];
    var { lcol } = TargetSheet[variant][0];
    var { lrow } = TargetSheet[variant][0];
    var { dataCol } = TargetSheet[variant][0];
    var { KudaCol } = TargetSheet[variant][0];
    var { kudaRow } = TargetSheet[variant][0];

    // ss -- is the current book
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var ReplaceCommas = getDecimalMarkIsCommaLocals(); // for Locals

    // [ 02 ]. If the sheet name is the same, you do business...
    var ls = ss.getSheetByName(LogSheet); // data sheet

    // [ 03 ]. Determine the level

    //-------------- The changing sheet --------------------------------
    var br = event.range;
    var scol = br.getColumn(); // the column number in which the change is made
    var srow = br.getRow(); // line number in which the change is made
    // Test if column fits
    if (scol >= lcol) {
      // Test if row fits
      if (srow >= lrow) {
        // adjust the level to size of
        // range that was changed
        var ColNum = br.getWidth();

        CurrentLevel = scol - lcol + ColNum + 1;

        // also need to adjust the range 'br'
        // split rows
        var RowNum = br.getHeight();
        if (ColNum > 1) {
          br = br.offset(0, ColNum - 1, RowNum, 1);
        } // wide range

        var HeadLevel = CurrentLevel - 1; // main level

        var X = NumOfLevels - CurrentLevel + 1;

        // the current level should not exceed the number of levels, or
        // we go beyond the desired range
        if (CurrentLevel <= NumOfLevels) {
          // determine columns on the sheet "Data"
          var KudaNado = ls.getRange(kudaRow, KudaCol);
          var lastRow = ls.getLastRow(); // get the address of the last cell
          var ChtoNado = ls.getRange(kudaRow, KudaCol, lastRow, KudaCol);

          // ============================================================================= > loop >

          var CurrLevelBase = CurrentLevel;
          for (var j = 1; j <= RowNum; j++) {
            CurrentLevel = CurrLevelBase; // refresh first val
            loop0: for (var k = 1; k <= X; k++) {
              HeadLevel = HeadLevel + k - 1; // adjust parent level
              var che = CurrentLevel + k - 1;

              CurrentLevel = CurrLevelBase + k - 1; // adjust current level

              var r = br.getCell(j, 1).offset(0, k - 1, 1);
              var SearchText = r.getValue(); // searched text
              // if anything is choosen!
              if (SearchText != '') {
                //-------------------------------------------------------------------

                // [ 04 ]. define variables to costumize data
                // for future data validation
                //--------------- Sheet with data --------------------------

                // values for check
                var checkVal = [];
                var checkDisplayVal = [];
                var Offs = CurrentLevel - 2;
                for (var s = Offs; s >= 0; s--) {
                  var checkR = r.offset(0, -s);
                  checkVal.push(checkR.getValue());
                  checkDisplayVal.push(checkR.getDisplayValue());
                }
                // get formula for validation
                var dataLevel = scol - lcol + ColNum + dataCol;
                var LookCol = colName(dataLevel - 1);

                var formula =
                  '=unique(filter(' + LookCol + '2:' + LookCol + lastRow;
                var Che = '';
                var Splinter = '';

                for (var i = 0; i < dataLevel - dataCol; i++) {
                  formula += FormulaSplitter;
                  LookCol = colName(dataCol - 1 + i);
                  formula += LookCol + '2:' + LookCol + lastRow;
                  Che = checkVal[i];
                  if (isNaN(Che)) {
                    Splinter = '"';
                  } else {
                    Splinter = '';

                    if (ReplaceCommas) {
                      // replace Dot(.) To Comma(,)

                      if (isNaN(checkDisplayVal[i])) {
                        Che = Che.toString();
                        Che = Che.replace('.', ',');
                      } else {
                        Splinter = '"';
                      }
                    }
                  }
                  if (typeof Che == 'number') {
                    formula += '=' + Che;
                  } else {
                    formula += '=' + Splinter + Che + Splinter;
                  }
                }
                formula += '))';
                KudaNado.setFormula(formula);

                var Response = [];

                loopP: for (var i = 1; i <= lastRow; i++) {
                  var currentValue = ChtoNado.getCell(i, 1).getValue();
                  if (currentValue != '') {
                    // Replace Dots to Cammas
                    if (ReplaceCommas) {
                      var CheckReplace = ChtoNado.getCell(
                        i,
                        1
                      ).getDisplayValue();
                      if (isNaN(currentValue) == false) {
                        if (isNaN(CheckReplace)) {
                          currentValue = currentValue.toString();
                          currentValue = currentValue.replace('.', ',');
                        }
                      }
                    }
                    Response.push(currentValue);
                  } else {
                    var Variants = i - 1; // number of possible values
                    break loopP; // exit loop
                  }
                }

                //-------------------------------------------------------------------

                // [ 05 ]. Build daya validation rule

                if (Variants == 0.0) {
                  break loop0;
                } else if (Variants >= 1.0) {
                  var cell = r.offset(0, 1);
                  var rule = SpreadsheetApp.newDataValidation()
                    .requireValueInList(Response, true)
                    .setAllowInvalid(false)
                    .build();
                  cell.setDataValidation(rule);
                }
                // other possibilities

                // TODO: ПОЧИНИТЬ АВТОЗАПОЛНЕНИЕ
                //Browser.msgBox(`CurrentLevel: ${CurrentLevel}`);

                //                if (Variants == 1.0)
                //                {
                //                  cell.setValue(Response[0]);
                //                  SearchText = null;
                //                  Response = null;
                //                  } // the only value
                //                 // break if many options
                if (Variants > 1) {
                  break loop0;
                }
              } // not blank cell
              else {
                // kill extra data validation if there were
                // columns on the right
                if (CurrentLevel <= NumOfLevels) {
                  for (var f = 1; f <= X; f++) {
                    var cell = r.offset(0, f);
                    // clean
                    cell.clear({ contentsOnly: true });
                    // get rid of validation
                    cell.clear({ validationsOnly: true });
                    // exit columns loop
                  }
                  break loop0;
                } // correct level
              } // empty row
            } // loop by cols
          } // loop by rows
          // ============================================================================= < loop <
        } // wrong level
      } // rows
    } // columns...
  } // main sheet
}

function onEdit(event) {
  SmartDataValidation(event);
}

function colName(n) {
  var ordA = 'a'.charCodeAt(0);
  var ordZ = 'z'.charCodeAt(0);

  var len = ordZ - ordA + 1;

  var s = '';
  while (n >= 0) {
    s = String.fromCharCode((n % len) + ordA) + s;
    n = Math.floor(n / len) - 1;
  }
  return s;
}

function getDecimalMarkIsCommaLocals() {
  // list of Locals Decimal mark = comma
  var LANGUAGE_BY_LOCALE = {
    af_NA: 'Afrikaans (Namibia)',
    af_ZA: 'Afrikaans (South Africa)',
    af: 'Afrikaans',
    sq_AL: 'Albanian (Albania)',
    sq: 'Albanian',
    ar_DZ: 'Arabic (Algeria)',
    ar_BH: 'Arabic (Bahrain)',
    ar_EG: 'Arabic (Egypt)',
    ar_IQ: 'Arabic (Iraq)',
    ar_JO: 'Arabic (Jordan)',
    ar_KW: 'Arabic (Kuwait)',
    ar_LB: 'Arabic (Lebanon)',
    ar_LY: 'Arabic (Libya)',
    ar_MA: 'Arabic (Morocco)',
    ar_OM: 'Arabic (Oman)',
    ar_QA: 'Arabic (Qatar)',
    ar_SA: 'Arabic (Saudi Arabia)',
    ar_SD: 'Arabic (Sudan)',
    ar_SY: 'Arabic (Syria)',
    ar_TN: 'Arabic (Tunisia)',
    ar_AE: 'Arabic (United Arab Emirates)',
    ar_YE: 'Arabic (Yemen)',
    ar: 'Arabic',
    hy_AM: 'Armenian (Armenia)',
    hy: 'Armenian',
    eu_ES: 'Basque (Spain)',
    eu: 'Basque',
    be_BY: 'Belarusian (Belarus)',
    be: 'Belarusian',
    bg_BG: 'Bulgarian (Bulgaria)',
    bg: 'Bulgarian',
    ca_ES: 'Catalan (Spain)',
    ca: 'Catalan',
    tzm_Latn: 'Central Morocco Tamazight (Latin)',
    tzm_Latn_MA: 'Central Morocco Tamazight (Latin, Morocco)',
    tzm: 'Central Morocco Tamazight',
    da_DK: 'Danish (Denmark)',
    da: 'Danish',
    nl_BE: 'Dutch (Belgium)',
    nl_NL: 'Dutch (Netherlands)',
    nl: 'Dutch',
    et_EE: 'Estonian (Estonia)',
    et: 'Estonian',
    fi_FI: 'Finnish (Finland)',
    fi: 'Finnish',
    fr_BE: 'French (Belgium)',
    fr_BJ: 'French (Benin)',
    fr_BF: 'French (Burkina Faso)',
    fr_BI: 'French (Burundi)',
    fr_CM: 'French (Cameroon)',
    fr_CA: 'French (Canada)',
    fr_CF: 'French (Central African Republic)',
    fr_TD: 'French (Chad)',
    fr_KM: 'French (Comoros)',
    fr_CG: 'French (Congo - Brazzaville)',
    fr_CD: 'French (Congo - Kinshasa)',
    fr_CI: 'French (Côte d’Ivoire)',
    fr_DJ: 'French (Djibouti)',
    fr_GQ: 'French (Equatorial Guinea)',
    fr_FR: 'French (France)',
    fr_GA: 'French (Gabon)',
    fr_GP: 'French (Guadeloupe)',
    fr_GN: 'French (Guinea)',
    fr_LU: 'French (Luxembourg)',
    fr_MG: 'French (Madagascar)',
    fr_ML: 'French (Mali)',
    fr_MQ: 'French (Martinique)',
    fr_MC: 'French (Monaco)',
    fr_NE: 'French (Niger)',
    fr_RW: 'French (Rwanda)',
    fr_RE: 'French (Réunion)',
    fr_BL: 'French (Saint Barthélemy)',
    fr_MF: 'French (Saint Martin)',
    fr_SN: 'French (Senegal)',
    fr_CH: 'French (Switzerland)',
    fr_TG: 'French (Togo)',
    fr: 'French',
    gl_ES: 'Galician (Spain)',
    gl: 'Galician',
    ka_GE: 'Georgian (Georgia)',
    ka: 'Georgian',
    de_AT: 'German (Austria)',
    de_BE: 'German (Belgium)',
    de_DE: 'German (Germany)',
    de_LI: 'German (Liechtenstein)',
    de_LU: 'German (Luxembourg)',
    de_CH: 'German (Switzerland)',
    de: 'German',
    el_CY: 'Greek (Cyprus)',
    el_GR: 'Greek (Greece)',
    el: 'Greek',
    hu_HU: 'Hungarian (Hungary)',
    hu: 'Hungarian',
    is_IS: 'Icelandic (Iceland)',
    is: 'Icelandic',
    id_ID: 'Indonesian (Indonesia)',
    id: 'Indonesian',
    it_IT: 'Italian (Italy)',
    it_CH: 'Italian (Switzerland)',
    it: 'Italian',
    kab_DZ: 'Kabyle (Algeria)',
    kab: 'Kabyle',
    kl_GL: 'Kalaallisut (Greenland)',
    kl: 'Kalaallisut',
    lv_LV: 'Latvian (Latvia)',
    lv: 'Latvian',
    lt_LT: 'Lithuanian (Lithuania)',
    lt: 'Lithuanian',
    mk_MK: 'Macedonian (Macedonia)',
    mk: 'Macedonian',
    naq_NA: 'Nama (Namibia)',
    naq: 'Nama',
    pl_PL: 'Polish (Poland)',
    pl: 'Polish',
    pt_BR: 'Portuguese (Brazil)',
    pt_GW: 'Portuguese (Guinea-Bissau)',
    pt_MZ: 'Portuguese (Mozambique)',
    pt_PT: 'Portuguese (Portugal)',
    pt: 'Portuguese',
    ro_MD: 'Romanian (Moldova)',
    ro_RO: 'Romanian (Romania)',
    ro: 'Romanian',
    ru_MD: 'Russian (Moldova)',
    ru_RU: 'Russian (Russia)',
    ru_UA: 'Russian (Ukraine)',
    ru: 'Russian',
    seh_MZ: 'Sena (Mozambique)',
    seh: 'Sena',
    sk_SK: 'Slovak (Slovakia)',
    sk: 'Slovak',
    sl_SI: 'Slovenian (Slovenia)',
    sl: 'Slovenian',
    es_AR: 'Spanish (Argentina)',
    es_BO: 'Spanish (Bolivia)',
    es_CL: 'Spanish (Chile)',
    es_CO: 'Spanish (Colombia)',
    es_CR: 'Spanish (Costa Rica)',
    es_DO: 'Spanish (Dominican Republic)',
    es_EC: 'Spanish (Ecuador)',
    es_SV: 'Spanish (El Salvador)',
    es_GQ: 'Spanish (Equatorial Guinea)',
    es_GT: 'Spanish (Guatemala)',
    es_HN: 'Spanish (Honduras)',
    es_419: 'Spanish (Latin America)',
    es_MX: 'Spanish (Mexico)',
    es_NI: 'Spanish (Nicaragua)',
    es_PA: 'Spanish (Panama)',
    es_PY: 'Spanish (Paraguay)',
    es_PE: 'Spanish (Peru)',
    es_PR: 'Spanish (Puerto Rico)',
    es_ES: 'Spanish (Spain)',
    es_US: 'Spanish (United States)',
    es_UY: 'Spanish (Uruguay)',
    es_VE: 'Spanish (Venezuela)',
    es: 'Spanish',
    sv_FI: 'Swedish (Finland)',
    sv_SE: 'Swedish (Sweden)',
    sv: 'Swedish',
    tr_TR: 'Turkish (Turkey)',
    tr: 'Turkish',
    uk_UA: 'Ukrainian (Ukraine)',
    uk: 'Ukrainian',
    vi_VN: 'Vietnamese (Vietnam)',
    vi: 'Vietnamese',
  };

  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var LocalS = SS.getSpreadsheetLocale();

  if (LANGUAGE_BY_LOCALE[LocalS] == undefined) {
    return false;
  }
  return true;
}

/*
function ReplaceDotsToCommas(dataIn) {
  var dataOut = dataIn.map(function(num) {
      if (isNaN(num)) {
        return num;
      }
      num = num.toString();
      return num.replace(".", ",");
  });
  return dataOut;
}

*/
