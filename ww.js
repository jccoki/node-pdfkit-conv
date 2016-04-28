/* vim: sw=2 ts=2 expandtab */
var express = require('express');
var app = express();

var PDFDocument = require('pdfkit');
var fs = require('fs');

// 8.27 x 11.69
// A4 standard size
var opts = {
  'page': {
    'cover': '/var/www/html/www.coredata.local/pdf-gen-2/img/cer/FrontPage3.png',
    'width': 595.44,
    'height': 841.68,
    'box_par_width': (612.28368 * 0.75),
    'margin': {
      'top': 72,
      'bottom': 72,
      'left': 72,
      'right': 72
    },
    'padding': {
      'top': 32,
      'bottom': 32,
      'left': 32,
      'right': 32,
    }
  },
  'font': {
    'weight': {
      'bold': 'node_modules/pdfkit/js/font/data/Arial_Bold.ttf',
      'italic': 'node_modules/pdfkit/js/font/data/Arial_Italic.ttf',
      'normal': 'node_modules/pdfkit/js/font/data/Arial.ttf'},
    'size': {
      'normal': 12,
      'h1': 14,
      'footer': 6}
    }
};

var output_dir = './Licensee'

app.get('/', function (req, res) {
  // start reading the XLSX file
  var XLSX = require('xlsx');
  var workbook = XLSX.readFile('WW\ 2014\ Consolidated\ Individual\ Report\ Data\ 150428\ \(Apogee\).xlsx');

  //excel file may contain multiple sheets, 1 sheet = 1 client
  var sheet_name_list = workbook.SheetNames;

  // get the range for rows and columns
  var ranges = workbook.Sheets[sheet_name_list[0]]['!ref'];
  var range = XLSX.utils.decode_range(ranges);
  var worksheet = workbook.Sheets[sheet_name_list[0]]

  // save the PDF properties variables
  var header_band_width = 8.64;
  var footer_band_width = 20.88;
  var pageOptions = {
    'size': [opts.page.width, opts.page.height],
    'margins': {'top': 0, 'left': 0, 'right': 0, 'bottom': 0}};

  var header_svg_path = 'M 0 0 l 0 ' + header_band_width + ' l ' + opts.page.width + ' 0 l 0 -' + header_band_width + ' l -'+ opts.page.width +' 0';

  var footer_svg_path = 'M 0 '+ (opts.page.height) +' l 0 -' + footer_band_width + ' l ' + opts.page.width + ' 0 l 0 ' + footer_band_width + ' l -'+ opts.page.width +' 0';

  // parsed at column 6
  var plannerName = '';

  // allow space between the table border and the bounding rect
  // -10pts for left margin, -7.5pts for right margin
  // set the table heading height
  var tableOpts = {
    'dividerWidth': 2,
    'margin': opts.page.margin.left + 15
  }

  var totalWidth = (opts.page.width-(opts.page.margin.left*2)) - (17.5 + (3*tableOpts['dividerWidth']));

  var ScoreAve = {}

  // Column 7
  ScoreAve['AquireAVE'] = {'planner': 50, 'mlc': 50, 'industry': 50}
  ScoreAve['Top1'] = {'mlc': 50, 'industry': 50}
  ScoreAve['Top2'] = {'mlc': 50, 'industry': 50}
  ScoreAve['Top3'] = {'mlc': 50, 'industry': 50}

  // assurance
  ScoreAve['AbilityDemo'] = {'mlc': 50, 'industry': 50}
  ScoreAve['AbilityExp'] = {'mlc': 50, 'industry': 50}
  ScoreAve['ClearEasy'] = {'mlc': 50, 'industry': 50}
  ScoreAve['AbDemoEff'] = {'mlc': 50, 'industry': 50}
  ScoreAve['ExperV'] = {'mlc': 50, 'industry': 50}
  ScoreAve['QualiV'] = {'mlc': 50, 'industry': 50}
  ScoreAve['ProdV'] = {'mlc': 50, 'industry': 50}
  ScoreAve['SrvcsV'] = {'mlc': 50, 'industry': 50}
  ScoreAve['Aa'] = {'mlc': 50, 'industry': 50}

  // compliance
  ScoreAve['RiskAtt'] = {'mlc': 50, 'industry': 50}
  ScoreAve['DisclPay'] = {'mlc': 50, 'industry': 50}
  ScoreAve['DisclFees'] = {'mlc': 50, 'industry': 50}
  ScoreAve['ShowFSG'] = {'mlc': 50, 'industry': 50}
  ScoreAve['ExplFSG'] = {'mlc': 50, 'industry': 50}
  ScoreAve['PrivyIssue'] = {'mlc': 50, 'industry': 50}
  ScoreAve['ExplIssue'] = {'mlc': 50, 'industry': 50}
  ScoreAve['SrvcProdOr'] = {'mlc': 50, 'industry': 50}
  ScoreAve['Cc'] = {'mlc': 50, 'industry': 50}

  // quality
  ScoreAve['ConvReco'] = {'mlc': 50, 'industry': 50}
  ScoreAve['FeesPay'] = {'mlc': 50, 'industry': 50}
  ScoreAve['PlanSrvcs'] = {'mlc': 50, 'industry': 50}
  ScoreAve['Qq'] = {'mlc': 50, 'industry': 50}

  // understanding
  ScoreAve['ListenSkill'] = {'mlc': 50, 'industry': 50}
  ScoreAve['Goals'] = {'mlc': 50, 'industry': 50}
  ScoreAve['DemoGoals'] = {'mlc': 50, 'industry': 50}
  ScoreAve['ReadFact'] = {'mlc': 50, 'industry': 50}
  ScoreAve['WellPrep'] = {'mlc': 50, 'industry': 50}
  ScoreAve['Uu'] = {'mlc': 50, 'industry': 50}

  // intention
  ScoreAve['M_2ndMeet'] = {'mlc': 50, 'industry': 50}
  ScoreAve['RecoP'] = {'mlc': 50, 'industry': 50}
  ScoreAve['Ii'] = {'mlc': 50, 'industry': 50}

  // reaction
  ScoreAve['Keen'] = {'mlc': 50, 'industry': 50}
  ScoreAve['Gimpress'] = {'mlc': 50, 'industry': 50}
  ScoreAve['Influence'] = {'mlc': 50, 'industry': 50}
  ScoreAve['Enthuse'] = {'mlc': 50, 'industry': 50}
  ScoreAve['Reltn'] = {'mlc': 50, 'industry': 50}
  ScoreAve['Rapprt'] = {'mlc': 50, 'industry': 50}
  ScoreAve['Probs'] = {'mlc': 50, 'industry': 50}
  ScoreAve['Honesty'] = {'mlc': 50, 'industry': 50}
  ScoreAve['Trust'] = {'mlc': 50, 'industry': 50}
  ScoreAve['Rr'] = {'mlc': 50, 'industry': 50}


  // environment
  ScoreAve['EasyTalk'] = {'mlc': 50, 'industry': 50}
  ScoreAve['SocCom'] = {'mlc': 50, 'industry': 50}
  ScoreAve['Friendly'] = {'mlc': 50, 'industry': 50}
  ScoreAve['OnTime'] = {'mlc': 50, 'industry': 50}
  ScoreAve['ProfDressV'] = {'mlc': 50, 'industry': 50}
  ScoreAve['StyleApp'] = {'mlc': 50, 'industry': 50}
  ScoreAve['LongAns'] = {'mlc': 50, 'industry': 50}
  ScoreAve['PeopSpeak'] = {'mlc': 50, 'industry': 50}
  ScoreAve['ContactFP'] = {'mlc': 50, 'industry': 50}
  ScoreAve['Helpful'] = {'mlc': 50, 'industry': 50}
  ScoreAve['EasyApp'] = {'mlc': 50, 'industry': 50}
  ScoreAve['ExtBldg'] = {'mlc': 50, 'industry': 50}
  ScoreAve['EnviBldg'] = {'mlc': 50, 'industry': 50}
  ScoreAve['Ee'] = {'mlc': 50, 'industry': 50}

  // follow up
  ScoreAve['FollowUp'] = {'mlc': 50, 'industry': 50}
  ScoreAve['DaysFollow'] = {'mlc': 50, 'industry': 50}
  ScoreAve['HowFollow'] = {'mlc': 50, 'industry': 50}

  var c_addr = '',
      question = '',
      answer = '',
      q_key = '',
      data = ''

  // variables for bounding rect box
  var box_height, box_width, box_par_width

  // read the row data one by one
  // start of usable row is 2, 0-index rule
  for(var Row = 2; Row <= range.e.r; Row++) {
    // fetch planner name
    plannerName = worksheet[ XLSX.utils.encode_cell({'c': 6, 'r': Row}) ].v

  var views_planner_offer_questions = []

    // mystery shopper profile questions
    mystery_shopper_profile_questions = [];
    for(var Col = 8; Col <= 10; Col++){
      c_addr = XLSX.utils.encode_cell({'c': Col, 'r': Row});
      // start of column = 8, end of column = 10, row = 1
      // strip " (Answer)"

      // questions are in 2nd row, answers are in current [row, column]
      question = worksheet[ XLSX.utils.encode_cell({'c': Col, 'r': 1}) ].v.replace(' (Answer)', '')

      // retrieving empty contents from cell address returns undefined
      // column 10 and 11 should be concatenated
      if(worksheet[c_addr]){
        if(Col == 10){
          answer = worksheet[c_addr].v + ' - ' + worksheet[XLSX.utils.encode_cell({'c': 11, 'r': Row})].v
        }else{
          answer = worksheet[c_addr].v
        }
      }else{
        answer = 'Irrelevant'
      }
      mystery_shopper_profile_questions.push({'question': question,'answer': answer})
    }

    // start of column = 12, end of column = 21, row = 1
    for(var Col = 12; Col <= 21; Col++){
      c_addr = XLSX.utils.encode_cell({'c': Col, 'r': Row});

      // retrieve question key
      q_key = worksheet[ XLSX.utils.encode_cell({'c': Col, 'r': 0}) ].v

      // retrieve question
      question = worksheet[ XLSX.utils.encode_cell({'c': Col, 'r': 1}) ].v.replace(' (Answer)', '')

      views_planner_offer_questions.push({'qkey': q_key, 'question': question, 'selfRate': worksheet[c_addr].v})
    }

    if( worksheet[ XLSX.utils.encode_cell({'c': 22, 'r': Row}) ] ){
      answer = worksheet[ XLSX.utils.encode_cell({'c': 22, 'r': Row}) ].v
    }else{
      answer = ''
    }
    var BenefFA = {
      'question': worksheet[ XLSX.utils.encode_cell({'c': 22, 'r': 1}) ].v.replace(' (Text)', ''),
      'answer': answer}

    var TopExpectations = []
    // start of column = 23, end of column = 25, row = 1
    for(var Col = 23; Col <= 25; Col++){
      TopExpectations.push({
        'key': worksheet[ XLSX.utils.encode_cell({'c': Col, 'r': 0}) ].v,
        'question': worksheet[ XLSX.utils.encode_cell({'c': Col, 'r': 1}) ].v,
        'selfRate': worksheet[ XLSX.utils.encode_cell({'c': Col, 'r': Row}) ].v,
      })
    }

    // start generating the reports one by one
    // start generating the pdf
    var doc = new PDFDocument(pageOptions)

    // Pipe it's output somewhere, like to a file or HTTP response
    // See below for browser usage
    doc.pipe(fs.createWriteStream(output_dir + '/Licensee_' + plannerName + '.pdf'))

    // header
    doc.path(header_svg_path).fillAndStroke("#febf38")
    doc.image('resources/coredata-logo.png', (opts.page.width*0.75), (header_band_width + (header_band_width*2)), {'scale': 0.80})
    // position the MLC logo 18% of the total page height and within 25% around the middle of the page
    doc.image('resources/ww-mlc-logo.png', ((opts.page.width/2)-(opts.page.width*.125)), (opts.page.height*0.18), {'scale': 0.35})

    // draw the middle band line starting from 33% of the total page height
    doc.path('M 0 '+ (opts.page.height*0.33) +' l 0 -' + header_band_width + ' l ' + opts.page.width + ' 0 l 0 ' + header_band_width + ' l -'+ opts.page.width +' 0').fillAndStroke("#c95109");

    // front page logo
    doc.image('resources/ww-frontpage-bg.png', 0, (opts.page.height*0.34), {'scale': 0.95})

    // draw the page title and its bounding rect box
    doc.path('M 0 '+ (opts.page.height) +' l 0 -' + (opts.page.height*0.30) + ' l ' + opts.page.width + ' 0 l 0 ' + (opts.page.height*0.30) + ' l -'+ opts.page.width +' 0').fillAndStroke("#c95109")
    doc.fontSize(22)
    doc.font(opts.font.weight.bold)
    doc.fillColor('#fff')
    // manually set the position of cursor for text alignment
    doc.text('Financial Planner Mystery Shopping', (opts.page.margin.left*1.25), (opts.page.height*0.75))

    lineHeight = doc.currentLineHeight()

    curr_x = doc.x
    curr_y = doc.y + lineHeight
    
    doc.fontSize(16)
    doc.font(opts.font.weight.normal)
    doc.text('Individual Planner Scorecard: ' + plannerName, curr_x, curr_y)

    // page 2
    doc.addPage(pageOptions)

    curr_x = opts.page.margin.left;
    curr_y = opts.page.margin.top;

    // footer and header
    doc.path(footer_svg_path).fillAndStroke("#c95109")
    doc.path(header_svg_path).fillAndStroke("#febf38")

    // title bounding rect
    // define header line width
    var box_width = opts.page.width-(opts.page.margin.left*2)
    // define header line height
    var box_height = opts.page.margin.top*0.6;

    doc.lineWidth(10)
    doc.lineJoin('round')
    doc.rect(curr_x, curr_y, box_width, box_height).fillAndStroke('#febf39').fillColor('#000')

    lineHeight = doc.currentLineHeight()
    curr_x = curr_x + 3;
    curr_y = curr_y + (lineHeight*0.60);

    doc.fontSize(13)
    doc.font(opts.font.weight.bold)
    doc.text('MLC/Garvan FP Mystery Shopping', curr_x, curr_y)
    doc.text('Individual Planner Scorecard – ' + plannerName)

    lineHeight = doc.currentLineHeight()
    curr_y = doc.y + (lineHeight*2);

    doc.fontSize(opts.font.size.normal)
    doc.text('Overall ACQUIRE Score', curr_x, curr_y)

    lineHeight = doc.currentLineHeight()
    curr_y = doc.y + (lineHeight*2);

    //set bounding rect paragraph width taking consideration of paragraph padding vs bounding rect
    box_par_width = box_width*0.95

    data = 'Your overall score was ' + ScoreAve['AquireAVE']['planner'] +' out of 100. The average score for MLC/Garvan FP financial planners was ' + ScoreAve['AquireAVE']['mlc'] + '. The average score for advisers in the broader Industry was ' + ScoreAve['AquireAVE']['industry'] + '.'

    doc.font(opts.font.weight.normal)

    doc.text(data.slice(0, 23), curr_x, (doc.y + lineHeight), {'width': box_par_width, 'align': 'left', 'continued': true})
    doc.font(opts.font.weight.bold)
    doc.text(data.slice(23, 25), {'continued': true})
    doc.font(opts.font.weight.normal)
    doc.text(data.slice(25, 97), {'continued': true})
    doc.font(opts.font.weight.bold)
    doc.text(data.slice(97, 99), {'continued': true})
    doc.font(opts.font.weight.normal)
    doc.text(data.slice(99, (data.length-3)), {'continued': true})
    doc.font(opts.font.weight.bold)
    doc.text(data.slice((data.length-3), (data.length-1)), {'continued': true})
    doc.font(opts.font.weight.normal)
    doc.text(data.slice((data.length-1), data.length))

    // the pink background
    lineHeight = doc.currentLineHeight()
    curr_y = doc.y + (lineHeight*3);

    // set current bounding rect width
    box_height = opts.page.height*0.37;
    box_par_width = box_width*0.90

    // draw a rounded bounding box
    doc.lineWidth(15)
    doc.lineJoin('round')
    doc.rect(curr_x, curr_y, box_width, box_height)
    doc.fillAndStroke('#ffddb1')

    curr_y = doc.y + (lineHeight*4);

    doc.fillColor('#000')
    doc.fontSize(opts.font.size.h1)
    doc.font(opts.font.weight.bold)
    doc.text('Mystery Shopper - Profile', doc.x, curr_y, {'align': 'center', 'width': box_width})

    lineHeight = doc.currentLineHeight()
    // explicitly align the current cursor to the previous paragraph text
    curr_x = opts.page.margin.left + 3;

    // iterate mystery_shopper_profile_questions
    doc.fontSize(opts.font.size.normal)
    for(var i = 0; i < mystery_shopper_profile_questions.length; i++){
      curr_y = doc.y + lineHeight;

      doc.text(mystery_shopper_profile_questions[i].question, curr_x, curr_y, {'width': box_par_width, 'align': 'left'});

      curr_y = doc.y + lineHeight;

      doc.font(opts.font.weight.normal)
      doc.text('Answer:', curr_x, curr_y)
      doc.font(opts.font.weight.bold)

      doc.text(mystery_shopper_profile_questions[i].answer, (curr_x+(box_par_width*0.15)), curr_y, {'width': box_par_width, 'align': 'left'})
    }

    // page 3
    doc.addPage(pageOptions)
    // footer and header
    doc.path(footer_svg_path).fillAndStroke("#c95109")
    doc.path(header_svg_path).fillAndStroke("#febf38")

    // draw the light blue background
    doc.lineWidth(20)
    doc.lineJoin('round')

    curr_x = opts.page.margin.left
    curr_y = opts.page.margin.top
    box_width = opts.page.width-(opts.page.margin.left*2)
    box_height = opts.page.height-(opts.page.margin.bottom*2.5)

    doc.rect(curr_x, curr_y, box_width, box_height)
    doc.fillAndStroke('#efeeed')
    doc.fillColor('#000')

    doc.fontSize(opts.font.size.h1)
    doc.font(opts.font.weight.bold)
    
    lineHeight = doc.currentLineHeight()

    curr_x = opts.page.margin.left + 5
    curr_y = opts.page.margin.top + lineHeight

    // this should move relative to the position of bounding box
    doc.text('Mystery Shopper – Views on Planner Offer', curr_x, curr_y, {'align': 'center', 'width': box_width});

    // this is also used by totalWidth
    box_par_width = opts.page.width-(opts.page.margin.left*2);

    // text position moves relative of the previous text
    doc.fontSize(opts.font.size.normal)
    lineHeight = doc.currentLineHeight()

    curr_x = opts.page.margin.left + 10
    curr_y = doc.y + (lineHeight*1.25)

    doc.text('How much do you think planners can assist you in the following:', curr_x, curr_y, {'width': box_par_width, 'align': 'justify'});

    // track current y-coord
    curr_y = doc.y + (lineHeight*1.25)

    doc.font(opts.font.weight.normal)

    // track the y-coords for table heading
    // the table headers
    var thead_heading = [];
    thead_heading[0] = {'width': 0.40, 'data': ''};
    thead_heading[1] = {'width': 0.20, 'data': 'Your rating'};
    thead_heading[2] = {'width': 0.20, 'data': 'MLC/Garvan FP average rating'};
    thead_heading[3] = {'width': 0.20, 'data': 'Industry average rating'};

    // resets back to pointy edge
    doc.lineWidth(1)
    // lineJoin() cause a subtle bug
    doc.lineCap('square');

    // track current width for the selected table column
    curr_x = tableOpts['margin'];
    curr_y = doc.y + (lineHeight*2);

    // draw the table headings
    doc.font(opts.font.weight.bold)
    doc.fontSize(9)

    lineHeight = doc.currentLineHeight()
    box_height = (lineHeight*2.25) + (lineHeight*0.4)
    for(var i=0; i<thead_heading.length; i++){
      doc.rect(curr_x, curr_y, (totalWidth*thead_heading[i].width), box_height)
      doc.fillAndStroke('#febf38')

      doc.fillColor('#000')
      doc.text(thead_heading[i].data, curr_x, (curr_y + (lineHeight*0.15)), {'align': 'center', 'width': (totalWidth*thead_heading[i].width)})

      curr_x = curr_x + (totalWidth*thead_heading[i].width) + tableOpts['dividerWidth']
    }

    doc.font(opts.font.weight.normal)

    // resets back to orig x-coord
    curr_x = tableOpts['margin'];

    // draw table rows
    // iterate the table row headings
    // add padding between rect box boundary and the text
    // each row is positioned using tableOpts['trow_height']*(i+1)
    lineHeight = doc.currentLineHeight()
    box_height = (lineHeight*2.25) + (lineHeight*0.4)
    for(var i = 0; i < views_planner_offer_questions.length; i++){
      var style = (i % 2) == 1 ? '#d9d9d9' : '#efeeed';

      // draw the bounding rect first
      curr_y = curr_y + box_height
      for(var j=0; j<thead_heading.length; j++){
        doc.rect(curr_x, curr_y, (totalWidth*thead_heading[j].width), box_height)
        doc.fillAndStroke(style)

        curr_x += (totalWidth*thead_heading[j].width) + tableOpts['dividerWidth']
      }

      // resets back to orig x-coord
      curr_x = tableOpts['margin']
      curr_y = curr_y + (lineHeight*0.15)

      // draw the text within the bounding rect
      doc.fillColor('#000')
      doc.text(views_planner_offer_questions[i].question, (curr_x + ((totalWidth*thead_heading[0].width)*0.015)), curr_y, {'width': ((totalWidth*thead_heading[0].width) - ((totalWidth*thead_heading[0].width)*0.03)), 'align': 'left'});

      // draw the self rating
      doc.font(opts.font.weight.bold)
      curr_x = curr_x+((totalWidth*thead_heading[0].width) + tableOpts['dividerWidth'])
      doc.text(views_planner_offer_questions[i].selfRate, curr_x, curr_y, {'width': (totalWidth*thead_heading[1].width), 'align': 'center'});
      doc.font(opts.font.weight.normal)

      // draw mcl/garvan rating
      curr_x = curr_x+((totalWidth*thead_heading[1].width) + tableOpts['dividerWidth'])
      doc.text(views_planner_offer_questions[i].selfRate, curr_x, curr_y, {'width': (totalWidth*thead_heading[2].width), 'align': 'center'});

      // draw industry average
      curr_x = curr_x+((totalWidth*thead_heading[2].width) + tableOpts['dividerWidth'])
      doc.text(views_planner_offer_questions[i].selfRate, curr_x, curr_y, {'width': (totalWidth*thead_heading[3].width), 'align': 'center'});

      // reset x-coord to the leftmost writable page of the page
      curr_x = tableOpts['margin'];

    }

    // draw BenefFA question code
    curr_x = opts.page.margin.left + 10
    curr_y = doc.y + (box_height*1.5);

    doc.font(opts.font.weight.bold)
    doc.fontSize(opts.font.size.normal)
    doc.text(BenefFA.question, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

    lineHeight = doc.currentLineHeight()
    doc.font(opts.font.weight.italic)
    doc.text(BenefFA.answer, curr_x, (doc.y + lineHeight), {'width': box_par_width, 'align': 'left'})

    curr_y = doc.y + (box_height*0.75);
    doc.font(opts.font.weight.bold)
    doc.text('What are the top three expectations you would have from a financial planner you are considering using (i.e. what do they have to do to get your business)?', curr_x, curr_y, {'width': box_par_width, 'align': 'left'});

    thead_heading = [];
    thead_heading[0] = {'width': 0.15, 'data': ''};
    thead_heading[1] = {'width': 0.35, 'data': 'Your shopper’s ranking'};
    thead_heading[2] = {'width': 0.30, 'data': 'MLC/Garvan FP top 3'};
    thead_heading[3] = {'width': 0.20, 'data': 'Industry top 3'};

    trow_heading = [];

    // assure correct sorting
    trow_heading[0] = 'Top 1';
    trow_heading[1] = 'Top 2';
    trow_heading[2] = 'Top 3';

    // track table coords
    doc.fontSize(9)
    doc.font(opts.font.weight.bold)
    lineHeight = doc.currentLineHeight()

    curr_x = tableOpts['margin'];
    curr_y = doc.y + (lineHeight*2)

    box_height = lineHeight*1.5

    // draw table heading
    for(var i=0; i<thead_heading.length; i++){
       doc.rect(curr_x, curr_y, (totalWidth*thead_heading[i].width), box_height)
       doc.fillAndStroke('#febf38')
       doc.fillColor('#000')
       doc.text(thead_heading[i].data, curr_x, (curr_y+(lineHeight*0.20)), {'align': 'center', 'width': (totalWidth*thead_heading[i].width)})

       curr_x += (totalWidth*thead_heading[i].width) + tableOpts['dividerWidth']
    }
    doc.font(opts.font.weight.normal)

    // draw table rows
    curr_y = curr_y + box_height

    // row level
    for(i=0; i<TopExpectations.length; i++){

      curr_x = tableOpts['margin'];
      // column level
      for(var j=0; j<thead_heading.length; j++){
        var style = (i % 2) == 1 ? '#d9d9d9' : '#efeeed';
        doc.rect(curr_x, curr_y, (totalWidth*thead_heading[j].width), box_height)
        doc.fillAndStroke(style)

        curr_x += (totalWidth*thead_heading[j].width) + tableOpts['dividerWidth']
      }

      // reset x-coords before drawing text
      curr_x = tableOpts['margin'] + 5;

      q_key = TopExpectations[i].key

      doc.fillColor('#000')
      curr_y = curr_y + (lineHeight*0.20)
      // draw the questions
      doc.text(TopExpectations[i].question, curr_x, curr_y, {'width': (totalWidth*thead_heading[0].width), 'align': 'left'});

      // draw the self rating

      doc.font(opts.font.weight.bold)
      curr_x = curr_x+((totalWidth*thead_heading[0].width) + tableOpts['dividerWidth'])
      doc.text(TopExpectations[i].selfRate, curr_x, curr_y, {'width': (totalWidth*thead_heading[1].width), 'align': 'center'});
      doc.font(opts.font.weight.normal)

      // draw mcl/garvan rating
      curr_x = curr_x+((totalWidth*thead_heading[1].width) + tableOpts['dividerWidth'])
      doc.text(ScoreAve[q_key]['mlc'], curr_x, curr_y, {'width': (totalWidth*thead_heading[2].width), 'align': 'center'});

      // draw industry average
      curr_x = curr_x+((totalWidth*thead_heading[2].width) + tableOpts['dividerWidth'])
      doc.text(ScoreAve[q_key]['industry'], curr_x, curr_y, {'width': (totalWidth*thead_heading[3].width), 'align': 'center'});

      // reset x-coord to the leftmost writable page of the page
      curr_x = tableOpts['margin'];

      curr_y = curr_y + box_height
    }

    // Assurance page
    doc.addPage(pageOptions)
    // footer and header
    doc.path(footer_svg_path).fillAndStroke("#c95109")
    doc.path(header_svg_path).fillAndStroke("#febf38")

    doc.lineWidth(20)
    doc.lineJoin('round')

    curr_x = opts.page.margin.left;
    curr_y = opts.page.margin.top;

    // standard opening page heading for the succedding pages
    box_width = opts.page.width-(opts.page.margin.left*2);
    box_height = opts.page.margin.top*0.4;
    doc.rect(curr_x, curr_y, box_width, box_height)
    doc.fillAndStroke('#febf38')
    doc.fillColor('#000')

    curr_x = opts.page.margin.left + 5;
    curr_y = opts.page.margin.top + 5;

    doc.font(opts.font.weight.bold)
    doc.fontSize(opts.font.size.normal)
    doc.text('Assurance', curr_x, curr_y, {'width': box_width, 'align': 'left'})

    box_par_width = box_width - 10

    doc.fontSize(11)
    doc.font(opts.font.weight.normal)
    doc.text('....Ability of planner to demonstrate and communicate knowledge/skills', curr_x, (curr_y + lineHeight), {'width': box_par_width, 'align': 'right'})

    lineHeight = doc.currentLineHeight()
    curr_y = doc.y + (lineHeight*3)

    doc.fontSize(opts.font.size.normal)

    lineHeight = doc.currentLineHeight()

    max_width = doc.widthOfString('MLC/Garvan FP Average: ') + 5

    // directly parse excel data here, Row is being iterated
    // CredPlanner
    question = worksheet[ XLSX.utils.encode_cell({'c': 26, 'r': 1}) ].v.replace(' (Text)', '')
    answer = worksheet[ XLSX.utils.encode_cell({'c': 26, 'r': Row}) ].v

    doc.font(opts.font.weight.bold)
    doc.text(question, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

    curr_y = doc.y + (lineHeight*1.25)
    doc.font(opts.font.weight.italic)
    doc.fontSize(11)
    doc.text(answer, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

    // CredPlanner
    question = worksheet[ XLSX.utils.encode_cell({'c': 27, 'r': 1}) ].v.replace(' (Text)', '')
    answer = worksheet[ XLSX.utils.encode_cell({'c': 27, 'r': Row}) ].v

    curr_y = doc.y + (lineHeight*1.25)
    doc.font(opts.font.weight.bold)
    doc.fontSize(opts.font.size.normal)
    doc.text(question, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

    curr_y = doc.y + (lineHeight*1.25)
    doc.font(opts.font.weight.italic)
    doc.fontSize(11)
    doc.text(answer, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

    curr_y = doc.y + (lineHeight*1.25)

    doc.fontSize(opts.font.size.normal)
    // 28 = AbilityDemo, 29 = AbilityExp 
    for(var i = 28; i <= 29; i++){
      q_key = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 0}) ].v
      question = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 1}) ].v
      question = question.replace(' (Value)', '').replace(' (Answer)', '').replace(' (Value)', '')

      answer = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': Row}) ].v

      doc.font(opts.font.weight.bold)
      doc.text(question, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

      curr_y = doc.y + (lineHeight*1.25)
      doc.font(opts.font.weight.normal)
      doc.text('Your Score:', curr_x, curr_y, {'align': 'left'})

      doc.font(opts.font.weight.bold)
      doc.text(answer, (curr_x + (max_width - 20) ), curr_y, {'align': 'left'})

      curr_y = doc.y
      doc.font(opts.font.weight.normal)
      doc.text('MLC/Garvan FP Average:   ' + ScoreAve[q_key]['mlc'] + '      |      Industry Average:   ' + ScoreAve[q_key]['industry'], curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

      curr_y = doc.y + (lineHeight*1.25)
    }

    // Impress
    question = worksheet[ XLSX.utils.encode_cell({'c': 30, 'r': 1}) ].v.replace(' (Text)', '')
    answer = worksheet[ XLSX.utils.encode_cell({'c': 30, 'r': Row}) ].v

    doc.font(opts.font.weight.bold)
    doc.text(question, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

    curr_y = doc.y + (lineHeight*1.25)
    doc.font(opts.font.weight.italic)
    doc.fontSize(11)
    doc.text(answer, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

    curr_y = doc.y + (lineHeight*1.25)

    doc.fontSize(opts.font.size.normal)
    doc.font(opts.font.weight.normal)
    // 31 = ClearEasy, 32 = AbDemoEff
    for(var i = 31; i <= 32; i++){
      q_key = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 0}) ].v
      question = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 1}) ].v
      question = question.replace(' (Value)', '').replace(' (Answer)', '').replace(' (Value)', '')

      answer = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': Row}) ].v

      doc.font(opts.font.weight.bold)
      doc.text(question, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

      curr_y = doc.y + (lineHeight*1.25)
      doc.font(opts.font.weight.normal)
      doc.text('Your Score:', curr_x, curr_y, {'align': 'left'})

      doc.font(opts.font.weight.bold)
      doc.text(answer, (curr_x + (max_width - 20) ), curr_y, {'align': 'left'})

      curr_y = doc.y
      doc.font(opts.font.weight.normal)
      doc.text('MLC/Garvan FP Average:   ' + ScoreAve[q_key]['mlc'] + '      |      Industry Average:   ' + ScoreAve[q_key]['industry'], curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

      curr_y = doc.y + (lineHeight*1.25)
    }


    // Assurance page 2
    doc.addPage(pageOptions)
    // footer and header
    doc.path(footer_svg_path).fillAndStroke("#c95109")
    doc.path(header_svg_path).fillAndStroke("#febf38")

    box_height = opts.page.margin.top*0.4;
    doc.fillColor('#000')

    lineHeight = doc.currentLineHeight()
    curr_y = doc.y + box_height + (lineHeight*2)

    doc.font(opts.font.weight.bold)
    doc.text('How much did the planner discuss the following during the meeting?', curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

    curr_y = doc.y + (lineHeight*1.25)

    // 33 = ExperV, 34 = QualiV, 35 = ProdV, 36 = SrvcsV
    doc.fontSize(opts.font.size.normal)
    for(var i=33; i<=36; i++){
      q_key = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 0}) ].v
      question = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 1}) ].v.replace(' (Answer)', '').replace(' (Text)', '')
      answer = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': Row}) ].v

      doc.font(opts.font.weight.bold)
      doc.text(question, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

      // track current y-coord
      curr_y = doc.y + (lineHeight*1.25)
      doc.font(opts.font.weight.normal)
      doc.text('Your Score:', curr_x, curr_y, {'align': 'left'})

      doc.font(opts.font.weight.bold)
      doc.text(answer, (curr_x + max_width), curr_y, {'align': 'left'})

      curr_y = doc.y
      doc.font(opts.font.weight.normal)
      doc.text('MLC/Garvan FP Average:', curr_x, curr_y, {'align': 'left'})
      doc.text(ScoreAve[q_key]['mlc'], (curr_x + max_width), curr_y, {'align': 'left'})

      curr_y = doc.y
      doc.text('Other responses:', curr_x, curr_y, {'align': 'left'})
      doc.text(answer, (curr_x + max_width), curr_y, {'align': 'left'})

      curr_y = doc.y
      doc.text('Industry Average:', curr_x, curr_y, {'align': 'left'})
      doc.text(ScoreAve[q_key]['industry'], (curr_x + max_width), curr_y, {'align': 'left'})

      curr_y = doc.y + (lineHeight*1.25)
    }

    // Aa
    curr_y = doc.y + (lineHeight*2)

    doc.font(opts.font.weight.bold)
    doc.text('Overall Assurance Score', curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

    lineHeight = doc.currentLineHeight()

    q_key = worksheet[ XLSX.utils.encode_cell({'c': 37, 'r': 0}) ].v.replace(' (Value)', '')
    answer = worksheet[ XLSX.utils.encode_cell({'c': 37, 'r': Row}) ].v
    doc.fontSize(opts.font.size.normal)

    max_width = doc.widthOfString('MLC/Garvan FP Assurances Average:') + 5

    curr_y = doc.y + (lineHeight*1.5)
    doc.font(opts.font.weight.bold)
    doc.text('Your Assurances Score:', curr_x, curr_y)
    doc.text(answer, (curr_x + (max_width - 40)), curr_y)

    doc.font(opts.font.weight.normal)
    
    curr_y = doc.y

    doc.text('MLC/Garvan FP Assurances Average:   ' + ScoreAve[q_key]['mlc'] + '      |      Industry Assurances Average:   ' + ScoreAve[q_key]['industry'], curr_x, curr_y, {'width': box_par_width, 'align': 'left'})


    // Compliance page
    doc.addPage(pageOptions)
    // footer and header
    doc.path(footer_svg_path).fillAndStroke("#c95109")
    doc.path(header_svg_path).fillAndStroke("#febf38")

    doc.lineWidth(20)
    doc.lineJoin('round')

    curr_x = opts.page.margin.left;
    curr_y = opts.page.margin.top;

    // standard opening page heading for the succedding pages
    box_width = opts.page.width-(opts.page.margin.left*2);
    box_height = opts.page.margin.top*0.4;
    doc.rect(curr_x, curr_y, box_width, box_height)
    doc.fillAndStroke('#febf38')
    doc.fillColor('#000')

    curr_x = opts.page.margin.left + 5;
    curr_y = opts.page.margin.top + 5;

    doc.font(opts.font.weight.bold)
    doc.fontSize(opts.font.size.normal)
    doc.text('Compliance', curr_x, curr_y, {'width': box_width, 'align': 'left'})

    box_par_width = box_width - 10

    doc.fontSize(11)
    doc.font(opts.font.weight.normal)
    doc.text('....Ability to satisfy relevant financial regulations', curr_x, (curr_y + lineHeight), {'width': box_par_width, 'align': 'right'})

    lineHeight = doc.currentLineHeight()
    curr_y = doc.y + (lineHeight*3)

    doc.fontSize(opts.font.size.normal)

    lineHeight = doc.currentLineHeight()

    max_width = doc.widthOfString('MLC/Garvan FP Average: ') + 5

    // 38 = RiskAtt, 39 = DisclPay, 40 = DisclFees, 41 = ShowFSG, 42 = ExplFSG
    for(var i=38; i<=42; i++){
      q_key = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 0}) ].v
      question = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 1}) ].v.replace(' (Answer)', '').replace(' (Text)', '')
      answer = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': Row}) ].v

      doc.font(opts.font.weight.bold)
      doc.text(question, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

      // track current y-coord
      curr_y = doc.y + (lineHeight*1.25)
      doc.font(opts.font.weight.normal)
      doc.text('Your Score:', curr_x, curr_y, {'align': 'left'})

      doc.font(opts.font.weight.bold)
      doc.text(answer, (curr_x + max_width), curr_y, {'align': 'left'})

      curr_y = doc.y
      doc.font(opts.font.weight.normal)
      doc.text('MLC/Garvan FP Average:', curr_x, curr_y, {'align': 'left'})
      doc.text(ScoreAve[q_key]['mlc'], (curr_x + max_width), curr_y, {'align': 'left'})

      curr_y = doc.y
      doc.text('Other responses:', curr_x, curr_y, {'align': 'left'})
      doc.text(answer, (curr_x + max_width), curr_y, {'align': 'left'})

      curr_y = doc.y
      doc.text('Industry Average:', curr_x, curr_y, {'align': 'left'})
      doc.text(ScoreAve[q_key]['industry'], (curr_x + max_width), curr_y, {'align': 'left'})

      curr_y = doc.y + (lineHeight*1.25)
    }

    // Compliance page 2
    doc.addPage(pageOptions)
    // footer and header
    doc.path(footer_svg_path).fillAndStroke("#c95109")
    doc.path(header_svg_path).fillAndStroke("#febf38")

    box_height = opts.page.margin.top*0.4;
    doc.fillColor('#000')

    lineHeight = doc.currentLineHeight()
    curr_y = doc.y + box_height + (lineHeight*2)

    // 43 = PrivyIssue, 44 = ExplIssue, 45 = SrvcProdOr
    for(var i=43; i<=45; i++){
      q_key = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 0}) ].v
      question = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 1}) ].v.replace(' (Answer)', '').replace(' (Text)', '')
      answer = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': Row}) ].v

      doc.font(opts.font.weight.bold)
      doc.text(question, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

      // track current y-coord
      curr_y = doc.y + (lineHeight*1.25)
      doc.font(opts.font.weight.normal)
      doc.text('Your Score:', curr_x, curr_y, {'align': 'left'})

      doc.font(opts.font.weight.bold)
      doc.text(answer, (curr_x + max_width), curr_y, {'align': 'left'})

      curr_y = doc.y
      doc.font(opts.font.weight.normal)
      doc.text('MLC/Garvan FP Average:', curr_x, curr_y, {'align': 'left'})
      doc.text(ScoreAve[q_key]['mlc'], (curr_x + max_width), curr_y, {'align': 'left'})

      curr_y = doc.y
      doc.text('Other responses:', curr_x, curr_y, {'align': 'left'})
      doc.text(answer, (curr_x + max_width), curr_y, {'align': 'left'})

      curr_y = doc.y
      doc.text('Industry Average:', curr_x, curr_y, {'align': 'left'})
      doc.text(ScoreAve[q_key]['industry'], (curr_x + max_width), curr_y, {'align': 'left'})

      curr_y = doc.y + (lineHeight*1.25)
    }

    // Cc
    curr_y = doc.y + (lineHeight*2)

    doc.font(opts.font.weight.bold)
    doc.text('Overall Compliance Score', curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

    lineHeight = doc.currentLineHeight()

    q_key = worksheet[ XLSX.utils.encode_cell({'c': 46, 'r': 0}) ].v.replace(' (Value)', '')
    answer = worksheet[ XLSX.utils.encode_cell({'c': 46, 'r': Row}) ].v
    doc.fontSize(opts.font.size.normal)

    max_width = doc.widthOfString('MLC/Garvan FP Compliance Average:') + 5

    curr_y = doc.y + (lineHeight*1.5)
    doc.font(opts.font.weight.bold)
    doc.text('Your Compliance Score:', curr_x, curr_y)
    doc.text(answer, (curr_x + (max_width - 40)), curr_y)

    doc.font(opts.font.weight.normal)
    
    curr_y = doc.y

    doc.text('MLC/Garvan FP Compliance Average:   ' + ScoreAve[q_key]['mlc'] + '      |      Industry Compliance Average:   ' + ScoreAve[q_key]['industry'], curr_x, curr_y, {'width': box_par_width, 'align': 'left'})


    // Quality page
    doc.addPage(pageOptions)
    // footer and header
    doc.path(footer_svg_path).fillAndStroke("#c95109")
    doc.path(header_svg_path).fillAndStroke("#febf38")

    doc.lineWidth(20)
    doc.lineJoin('round')

    curr_x = opts.page.margin.left;
    curr_y = opts.page.margin.top;

    // standard opening page heading for the succedding pages
    box_width = opts.page.width-(opts.page.margin.left*2);
    box_height = opts.page.margin.top*0.4;
    doc.rect(curr_x, curr_y, box_width, box_height)
    doc.fillAndStroke('#febf38')
    doc.fillColor('#000')

    curr_x = opts.page.margin.left + 5;
    curr_y = opts.page.margin.top + 5;

    doc.font(opts.font.weight.bold)
    doc.fontSize(opts.font.size.normal)
    doc.text('Quality', curr_x, curr_y, {'width': box_width, 'align': 'left'})

    box_par_width = box_width - 10

    doc.fontSize(11)
    doc.font(opts.font.weight.normal)
    doc.text('....Ability to satisfy customer needs and provide perceived value', curr_x, (curr_y + lineHeight), {'width': box_par_width, 'align': 'right'})

    lineHeight = doc.currentLineHeight()
    curr_y = doc.y + (lineHeight*3)

    doc.fontSize(opts.font.size.normal)

    lineHeight = doc.currentLineHeight()

    max_width = doc.widthOfString('MLC/Garvan FP Average: ') + 5

    // 47 = ConvReco, 48 = FeesPay, 49 = PlanSrvcs
    for(var i = 47; i <= 49; i++){
      q_key = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 0}) ].v
      question = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 1}) ].v
      question = question.replace(' (Value)', '').replace(' (Answer)', '').replace(' (Value)', '')

      answer = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': Row}) ].v

      doc.font(opts.font.weight.bold)
      doc.text(question, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

      curr_y = doc.y + (lineHeight*1.25)
      doc.font(opts.font.weight.normal)
      doc.text('Your Score:', curr_x, curr_y, {'align': 'left'})

      doc.font(opts.font.weight.bold)
      doc.text(answer, (curr_x + (max_width - 20) ), curr_y, {'align': 'left'})

      curr_y = doc.y
      doc.font(opts.font.weight.normal)
      doc.text('MLC/Garvan FP Average:   ' + ScoreAve[q_key]['mlc'] + '      |      Industry Average:   ' + ScoreAve[q_key]['industry'], curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

      curr_y = doc.y + (lineHeight*1.25)
    }

    // Qq
    curr_y = doc.y + (lineHeight*2)

    doc.font(opts.font.weight.bold)
    doc.text('Overall Quality Score', curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

    lineHeight = doc.currentLineHeight()

    q_key = worksheet[ XLSX.utils.encode_cell({'c': 79, 'r': 0}) ].v.replace(' (Value)', '')
    answer = worksheet[ XLSX.utils.encode_cell({'c': 79, 'r': Row}) ].v
    doc.fontSize(opts.font.size.normal)

    max_width = doc.widthOfString('MLC/Garvan FP Quality Average:') + 5

    curr_y = doc.y + (lineHeight*1.5)
    doc.font(opts.font.weight.bold)
    doc.text('Your Quality Score:', curr_x, curr_y)
    doc.text(answer, (curr_x + (max_width - 40)), curr_y)

    doc.font(opts.font.weight.normal)
    
    curr_y = doc.y

    doc.text('MLC/Garvan FP Quality Average:   ' + ScoreAve[q_key]['mlc'] + '      |      Industry Quality Average:   ' + ScoreAve[q_key]['industry'], curr_x, curr_y, {'width': box_par_width, 'align': 'left'})


    // Understanding page
    doc.addPage(pageOptions)
    // footer and header
    doc.path(footer_svg_path).fillAndStroke("#c95109")
    doc.path(header_svg_path).fillAndStroke("#febf38")

    doc.lineWidth(20)
    doc.lineJoin('round')

    curr_x = opts.page.margin.left;
    curr_y = opts.page.margin.top;

    // standard opening page heading for the succedding pages
    box_width = opts.page.width-(opts.page.margin.left*2);
    box_height = opts.page.margin.top*0.4;
    doc.rect(curr_x, curr_y, box_width, box_height)
    doc.fillAndStroke('#febf38')
    doc.fillColor('#000')

    curr_x = opts.page.margin.left + 5;
    curr_y = opts.page.margin.top + 5;

    doc.font(opts.font.weight.bold)
    doc.fontSize(opts.font.size.normal)
    doc.text('Understanding', curr_x, curr_y, {'width': box_width, 'align': 'left'})

    box_par_width = box_width - 10

    doc.fontSize(11)
    doc.font(opts.font.weight.normal)
    doc.text('....Ability to understand client needs', curr_x, (curr_y + lineHeight), {'width': box_par_width, 'align': 'right'})

    lineHeight = doc.currentLineHeight()
    curr_y = doc.y + (lineHeight*3)

    doc.fontSize(opts.font.size.normal)

    lineHeight = doc.currentLineHeight()

    max_width = doc.widthOfString('MLC/Garvan FP Average: ') + 5

    // 51 = ListenSkill, 52 = Goals, 53 = DemoGoals, 54 = ReadFact, 55 = WellPrep
    for(var i = 51; i <= 55; i++){
      q_key = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 0}) ].v
      question = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 1}) ].v
      question = question.replace(' (Value)', '').replace(' (Answer)', '').replace(' (Value)', '')

      answer = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': Row}) ].v

      doc.font(opts.font.weight.bold)
      doc.text(question, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

      curr_y = doc.y + (lineHeight*1.25)
      doc.font(opts.font.weight.normal)
      doc.text('Your Score:', curr_x, curr_y, {'align': 'left'})

      doc.font(opts.font.weight.bold)
      doc.text(answer, (curr_x + (max_width - 20) ), curr_y, {'align': 'left'})

      curr_y = doc.y
      doc.font(opts.font.weight.normal)
      doc.text('MLC/Garvan FP Average:   ' + ScoreAve[q_key]['mlc'] + '      |      Industry Average:   ' + ScoreAve[q_key]['industry'], curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

      curr_y = doc.y + (lineHeight*1.25)
    }


    // Uu
    curr_y = doc.y + (lineHeight*2)

    doc.font(opts.font.weight.bold)
    doc.text('Overall Understanding Score', curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

    lineHeight = doc.currentLineHeight()

    q_key = worksheet[ XLSX.utils.encode_cell({'c': 79, 'r': 0}) ].v.replace(' (Value)', '')
    answer = worksheet[ XLSX.utils.encode_cell({'c': 79, 'r': Row}) ].v
    doc.fontSize(opts.font.size.normal)

    max_width = doc.widthOfString('MLC/Garvan FP Understanding Average:') + 5

    curr_y = doc.y + (lineHeight*1.5)
    doc.font(opts.font.weight.bold)
    doc.text('Your Understanding Score:', curr_x, curr_y)
    doc.text(answer, (curr_x + (max_width - 40)), curr_y)

    doc.font(opts.font.weight.normal)
    
    curr_y = doc.y

    doc.text('MLC/Garvan FP Understanding Average:   ' + ScoreAve[q_key]['mlc'] + '      |      Industry Understanding Average:   ' + ScoreAve[q_key]['industry'], curr_x, curr_y, {'width': box_par_width, 'align': 'left'})


    // Intention page
    doc.addPage(pageOptions)
    // footer and header
    doc.path(footer_svg_path).fillAndStroke("#c95109")
    doc.path(header_svg_path).fillAndStroke("#febf38")

    doc.lineWidth(20)
    doc.lineJoin('round')

    curr_x = opts.page.margin.left;
    curr_y = opts.page.margin.top;

    // standard opening page heading for the succedding pages
    box_width = opts.page.width-(opts.page.margin.left*2);
    box_height = opts.page.margin.top*0.4;
    doc.rect(curr_x, curr_y, box_width, box_height)
    doc.fillAndStroke('#febf38')
    doc.fillColor('#000')

    curr_x = opts.page.margin.left + 5;
    curr_y = opts.page.margin.top + 5;

    doc.font(opts.font.weight.bold)
    doc.fontSize(opts.font.size.normal)
    doc.text('Intention', curr_x, curr_y, {'width': box_width, 'align': 'left'})

    box_par_width = box_width - 10

    doc.fontSize(11)
    doc.font(opts.font.weight.normal)
    doc.text('....Client intention to use/reuse/recommend planner', curr_x, (curr_y + lineHeight), {'width': box_par_width, 'align': 'right'})

    lineHeight = doc.currentLineHeight()
    curr_y = doc.y + (lineHeight*3)

    doc.fontSize(opts.font.size.normal)

    lineHeight = doc.currentLineHeight()

    max_width = doc.widthOfString('MLC/Garvan FP Average: ') + 5

    for(var i=57; i<=60; i=i+2){
      q_key = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 0}) ].v
      question = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 1}) ].v
      question = question.replace(' (Value)', '').replace(' (Answer)', '')

      answer = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': Row}) ].v

      doc.font(opts.font.weight.bold)
      doc.text(question, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

      curr_y = doc.y + lineHeight
      doc.font(opts.font.weight.normal)
      doc.text('Your Score:', curr_x, curr_y, {'align': 'left'})

      doc.font(opts.font.weight.bold)
      doc.text(answer, (curr_x + (max_width - 40)), curr_y, {'align': 'left'})

      curr_y = doc.y
      doc.font(opts.font.weight.normal)
      doc.text('MLC/Garvan FP Average:   ' + ScoreAve[q_key]['mlc'] + '      |      Industry Average:   ' + ScoreAve[q_key]['industry'], curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

      answer = worksheet[ XLSX.utils.encode_cell({'c': (i+1), 'r': Row}) ].v

      curr_y = doc.y + (lineHeight*1.25)
      doc.font(opts.font.weight.bold)
      doc.text('Reason Given:', curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

      curr_y = doc.y + (lineHeight*1.25)
      doc.font(opts.font.weight.normal)
 
      doc.text(answer, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

      curr_y = doc.y + (lineHeight*1.25)
    }

    // Ii
    curr_y = doc.y + (lineHeight*2)

    doc.font(opts.font.weight.bold)
    doc.text('Overall Intention Score', curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

    lineHeight = doc.currentLineHeight()

    q_key = worksheet[ XLSX.utils.encode_cell({'c': 79, 'r': 0}) ].v.replace(' (Value)', '')
    answer = worksheet[ XLSX.utils.encode_cell({'c': 79, 'r': Row}) ].v
    doc.fontSize(opts.font.size.normal)

    max_width = doc.widthOfString('MLC/Garvan FP Intention Average:') + 5

    curr_y = doc.y + (lineHeight*1.5)
    doc.font(opts.font.weight.bold)
    doc.text('Your Intention Score', curr_x, curr_y)
    doc.text(answer, (curr_x + (max_width - 40)), curr_y)

    doc.font(opts.font.weight.normal)
    
    curr_y = doc.y

    doc.text('MLC/Garvan FP Intention Average:   ' + ScoreAve[q_key]['mlc'] + '      |      Industry Intention Average:   ' + ScoreAve[q_key]['industry'], curr_x, curr_y, {'width': box_par_width, 'align': 'left'})


    // Reaction page
    doc.addPage(pageOptions)
    // footer and header
    doc.path(footer_svg_path).fillAndStroke("#c95109")
    doc.path(header_svg_path).fillAndStroke("#febf38")

    doc.lineWidth(20)
    doc.lineJoin('round')

    curr_x = opts.page.margin.left;
    curr_y = opts.page.margin.top;

    // standard opening page heading for the succedding pages
    box_width = opts.page.width-(opts.page.margin.left*2);
    box_height = opts.page.margin.top*0.4;
    doc.rect(curr_x, curr_y, box_width, box_height)
    doc.fillAndStroke('#febf38')
    doc.fillColor('#000')

    curr_x = opts.page.margin.left + 5;
    curr_y = opts.page.margin.top + 5;

    doc.font(opts.font.weight.bold)
    doc.fontSize(opts.font.size.normal)
    doc.text('Reaction', curr_x, curr_y, {'width': box_width, 'align': 'left'})

    box_par_width = box_width - 10

    doc.fontSize(11)
    doc.font(opts.font.weight.normal)
    doc.text('....Client’s emotive/affective response to purchase process', curr_x, (curr_y + lineHeight), {'width': box_par_width, 'align': 'right'})

    lineHeight = doc.currentLineHeight()
    curr_y = doc.y + (lineHeight*3)

    doc.fontSize(opts.font.size.normal)

    lineHeight = doc.currentLineHeight()

    max_width = doc.widthOfString('MLC/Garvan FP Average: ') + 5

    // Keen
    q_key = worksheet[ XLSX.utils.encode_cell({'c': 66, 'r': 0}) ].v
    question = worksheet[ XLSX.utils.encode_cell({'c': 66, 'r': 1}) ].v
    question = question.replace(' (Value)', '').replace(' (Answer)', '')

    answer = worksheet[ XLSX.utils.encode_cell({'c': 66, 'r': Row}) ].v

    doc.font(opts.font.weight.bold)
    doc.text(question, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

    curr_y = doc.y + (lineHeight*1.25)
    doc.font(opts.font.weight.normal)
    doc.text('Your Score:', curr_x, curr_y, {'align': 'left'})

    doc.font(opts.font.weight.bold)
    doc.text(answer, (curr_x + (max_width - 40)), curr_y, {'align': 'left'})

    curr_y = doc.y
    doc.font(opts.font.weight.normal)
    doc.text('MLC/Garvan FP Average:   ' + ScoreAve[q_key]['mlc'] + '      |      Industry Average:   ' + ScoreAve[q_key]['industry'], curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

    // KeenV
    if( worksheet[ XLSX.utils.encode_cell({'c': 67, 'r': Row}) ] ){
      answer = worksheet[ XLSX.utils.encode_cell({'c': 67, 'r': Row}) ].v
    }else{
      answer = ''
    }

    curr_y = doc.y + (lineHeight*1.25)
    doc.font(opts.font.weight.bold)
    doc.text('Please explain how the planner demonstrated their keenness for your business?', curr_x, curr_y, {'width': box_width, 'align': 'left'})

    curr_y = doc.y + (lineHeight*1.25)
    doc.font(opts.font.weight.normal)
    doc.text(answer, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

    curr_y = doc.y + (lineHeight*1.25)
    doc.font(opts.font.weight.bold)
    doc.text('How would you rate the engagement skills of the planner?', curr_x, curr_y, {'width': box_width, 'align': 'left'})

    curr_y = doc.y + (lineHeight*1.25)

    // 68 = Gimpress, 69 = Influence, 70 = Enthuse, 71 = Reltn, 72 = Rapprt
    for(var i=68; i<= 72; i++){
      q_key = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 0}) ].v
      question = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 1}) ].v
      question = question.replace(' (Value)', '').replace(' (Answer)', '')

      answer = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': Row}) ].v

      doc.font(opts.font.weight.bold)
      doc.text(question, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

      curr_y = doc.y + lineHeight
      doc.font(opts.font.weight.normal)
      doc.text('Your Score:', curr_x, curr_y, {'align': 'left'})

      doc.font(opts.font.weight.bold)
      doc.text(answer, (curr_x + (max_width - 40)), curr_y, {'align': 'left'})

      curr_y = doc.y
      doc.font(opts.font.weight.normal)
      doc.text('MLC/Garvan FP Average:   ' + ScoreAve[q_key]['mlc'] + '      |      Industry Average:   ' + ScoreAve[q_key]['industry'], curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

      curr_y = doc.y + (lineHeight*1.25)
    }

    // Reaction page 2
    doc.addPage(pageOptions)
    // footer and header
    doc.path(footer_svg_path).fillAndStroke("#c95109")
    doc.path(header_svg_path).fillAndStroke("#febf38")

    box_height = opts.page.margin.top*0.4;
    doc.fillColor('#000')

    curr_y = doc.y + box_height + (lineHeight*2)

    doc.fontSize(opts.font.size.normal)
    lineHeight = doc.currentLineHeight()

    // 73 = Probs, 75 = Honesty, 77 = Trust
    for(var i=73; i<= 77; i=i+2){
      q_key = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 0}) ].v
      question = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 1}) ].v.replace(' (Answer)', '')
      answer = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': Row}) ].v

      doc.font(opts.font.weight.bold)
      doc.text(question, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

      //get max width for the questions
      doc.font(opts.font.weight.normal)
      max_width = doc.widthOfString('MLC/Garvan FP Average: ') + 5

      // track current y-coord
      curr_y = doc.y + (lineHeight*1.25)
      doc.text('Your Score:', curr_x, curr_y, {'width': max_width, 'align': 'left'})
      doc.font(opts.font.weight.bold)
      doc.text(answer, (doc.x + max_width), curr_y, {'width': (box_width*0.66), 'align': 'left'})

      curr_y = doc.y
      doc.font(opts.font.weight.normal)
      doc.text('MLC/Garvan FP Average:', curr_x, curr_y, {'width': max_width, 'align': 'left'})
      doc.text(ScoreAve[q_key]['mlc'], (doc.x + max_width), curr_y, {'width': (box_width*0.66), 'align': 'left'})

      curr_y = doc.y
      doc.text('Industry Average:', curr_x, (curr_y), {'width': max_width, 'align': 'left'})
      doc.text(ScoreAve[q_key]['industry'], (doc.x + max_width), curr_y, {'width': (box_width*0.66), 'align': 'left'})

      // text explanation
      c_addr = XLSX.utils.encode_cell({'c': i+1, 'r': Row})
      if(worksheet[c_addr]){
        answer = worksheet[c_addr].v
      }else{
        answer = 'No answer'
      }
      question = worksheet[ XLSX.utils.encode_cell({'c': (i+1), 'r': 1}) ].v.replace(' (Text)', '')

      curr_y = doc.y + (lineHeight*1.25)

      doc.font(opts.font.weight.bold)
      doc.text(question, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

      curr_y = doc.y + (lineHeight*1.25)

      doc.font(opts.font.weight.normal)
      doc.text(answer, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

      curr_y = doc.y + (lineHeight*1.25)
    }

    // Rr
    curr_y = doc.y + (lineHeight*2)

    doc.font(opts.font.weight.bold)
    doc.text('Overall Reaction Score', curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

    lineHeight = doc.currentLineHeight()

    q_key = worksheet[ XLSX.utils.encode_cell({'c': 79, 'r': 0}) ].v.replace(' (Value)', '')
    answer = worksheet[ XLSX.utils.encode_cell({'c': 79, 'r': Row}) ].v
    doc.fontSize(opts.font.size.normal)

    max_width = doc.widthOfString('MLC/Garvan FP Reaction Average:') + 5

    curr_y = doc.y + (lineHeight*1.5)
    doc.font(opts.font.weight.bold)
    doc.text('Your Reaction Score:', curr_x, curr_y)
    doc.text(answer, (curr_x + (max_width - 40)), curr_y)

    doc.font(opts.font.weight.normal)
    
    curr_y = doc.y

    doc.text('MLC/Garvan FP Reaction Average:   ' + ScoreAve[q_key]['mlc'] + '      |      Industry Average:   ' + ScoreAve[q_key]['industry'], curr_x, curr_y, {'width': box_par_width, 'align': 'left'})


    // Environment page
    doc.addPage(pageOptions)
    // footer and header
    doc.path(footer_svg_path).fillAndStroke("#c95109")
    doc.path(header_svg_path).fillAndStroke("#febf38")

    doc.lineWidth(20)
    doc.lineJoin('round')

    curr_x = opts.page.margin.left;
    curr_y = opts.page.margin.top;

    // standard opening page heading for the succedding pages
    box_width = opts.page.width-(opts.page.margin.left*2);
    box_height = opts.page.margin.top*0.4;
    doc.rect(curr_x, curr_y, box_width, box_height)
    doc.fillAndStroke('#febf38')
    doc.fillColor('#000')

    curr_x = opts.page.margin.left + 5;
    curr_y = opts.page.margin.top + 5;

    doc.font(opts.font.weight.bold)
    doc.fontSize(opts.font.size.normal)
    doc.text('Environment', curr_x, curr_y, {'width': box_width, 'align': 'left'})

    box_par_width = box_width - 10

    doc.fontSize(11)
    doc.font(opts.font.weight.normal)
    doc.text('....Intangible/tangible aspects of the client-adviser experience', curr_x, (curr_y + lineHeight), {'width': box_par_width, 'align': 'right'})

    lineHeight = doc.currentLineHeight()
    curr_y = doc.y + (lineHeight*3)

    doc.fontSize(opts.font.size.normal)

    lineHeight = doc.currentLineHeight()

    //get max width for the questions
    max_width = doc.widthOfString('MLC/Garvan FP Average: ') + 5

    curr_y = doc.y + box_height + (lineHeight*1.5)

    doc.font(opts.font.weight.bold)
    doc.text('How did the planner rate with regard to the following:', curr_x, curr_y, {'width': box_width, 'align': 'left'})

    curr_y = doc.y + lineHeight

    // 80 = EasyTalk, 81 = SocCom, 82 = Friendly, 83 = OnTime, 84 = ProfDressV, 85 = StyleApp
    for(var i=80; i<= 85; i++){
      q_key = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 0}) ].v
      question = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 1}) ].v
      question = question.replace(' (Value)', '').replace(' (Answer)', '').replace(' (Value)', '')

      answer = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': Row}) ].v

      doc.font(opts.font.weight.bold)
      doc.text(question, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

      curr_y = doc.y + (lineHeight*1.25)
      doc.font(opts.font.weight.normal)
      doc.text('Your Score:', curr_x, curr_y, {'align': 'left'})

      doc.font(opts.font.weight.bold)
      doc.text(answer, (curr_x + (max_width - 20) ), curr_y, {'align': 'left'})

      curr_y = doc.y
      doc.font(opts.font.weight.normal)
      doc.text('MLC/Garvan FP Average:   ' + ScoreAve[q_key]['mlc'] + '      |      Industry Average:   ' + ScoreAve[q_key]['industry'], curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

      curr_y = doc.y + (lineHeight*1.25)
    }

    // LongAns
    q_key = worksheet[ XLSX.utils.encode_cell({'c': 86, 'r': 0}) ].v
    question = worksheet[ XLSX.utils.encode_cell({'c': 86, 'r': 1}) ].v.replace(' (Answer)', '')
    answer = worksheet[ XLSX.utils.encode_cell({'c': 86, 'r': Row}) ].v

    doc.font(opts.font.weight.bold)
    doc.text(question, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

    curr_y = doc.y + (lineHeight*1.25)
    doc.font(opts.font.weight.normal)
    doc.text('Your Score:', curr_x, curr_y, {'align': 'left'})
    doc.text(answer, (doc.x + max_width), curr_y, {'align': 'left'})

    curr_y = doc.y
    doc.text('MLC/Garvan FP Average:', curr_x, curr_y, {'align': 'left'})
    doc.text(ScoreAve[q_key]['mlc'], (curr_x + max_width), curr_y, {'align': 'left'})

    curr_y = doc.y
    doc.text('Other responses:', curr_x, curr_y, {'align': 'left'})
    doc.text(answer, (curr_x + max_width), curr_y, {'align': 'left'})

    curr_y = doc.y
    doc.text('Industry Average:', curr_x, curr_y, {'align': 'left'})
    doc.text(ScoreAve[q_key]['industry'], (curr_x + max_width), curr_y, {'align': 'left'})

    // Environment page 2
    doc.addPage(pageOptions)
    // footer and header
    doc.path(footer_svg_path).fillAndStroke("#c95109")
    doc.path(header_svg_path).fillAndStroke("#febf38")

    box_height = opts.page.margin.top*0.4;
    doc.fillColor('#000')

    lineHeight = doc.currentLineHeight()
    curr_y = doc.y + box_height + (lineHeight*2)

    // 87 = PeopSpeak, 88 = ContactFP
    for(var i=87; i<= 88; i++){
      q_key = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 0}) ].v
      question = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 1}) ].v.replace(' (Answer)', '').replace(' (Text)', '')
      answer = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': Row}) ].v

      doc.font(opts.font.weight.bold)
      doc.text(question, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

      // track current y-coord
      curr_y = doc.y + (lineHeight*1.25)
      doc.font(opts.font.weight.normal)
      doc.text('Your Score:', curr_x, curr_y, {'align': 'left'})

      doc.font(opts.font.weight.bold)
      doc.text(answer, (curr_x + max_width), curr_y, {'align': 'left'})

      curr_y = doc.y
      doc.font(opts.font.weight.normal)
      doc.text('MLC/Garvan FP Average:', curr_x, curr_y, {'align': 'left'})
      doc.text(ScoreAve[q_key]['mlc'], (curr_x + max_width), curr_y, {'align': 'left'})

      curr_y = doc.y
      doc.text('Other responses:', curr_x, curr_y, {'align': 'left'})
      doc.text(answer, (curr_x + max_width), curr_y, {'align': 'left'})

      curr_y = doc.y
      doc.text('Industry Average:', curr_x, curr_y, {'align': 'left'})
      doc.text(ScoreAve[q_key]['industry'], (curr_x + max_width), curr_y, {'align': 'left'})

      curr_y = doc.y + (lineHeight*1.25)
    }

    // 89 = Helpful, 90 = EasyApp
    for(var i=89; i<= 90; i++){
      q_key = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 0}) ].v
      question = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 1}) ].v
      question = question.replace(' (Value)', '').replace(' (Answer)', '').replace(' (Value)', '')

      answer = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': Row}) ].v

      doc.font(opts.font.weight.bold)
      doc.text(question, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

      curr_y = doc.y + (lineHeight*1.25)
      doc.font(opts.font.weight.normal)
      doc.text('Your Score:', curr_x, curr_y, {'align': 'left'})

      doc.font(opts.font.weight.bold)
      doc.text(answer, (curr_x + (max_width - 20) ), curr_y, {'align': 'left'})

      curr_y = doc.y
      doc.font(opts.font.weight.normal)
      doc.text('MLC/Garvan FP Average:   ' + ScoreAve[q_key]['mlc'] + '      |      Industry Average:   ' + ScoreAve[q_key]['industry'], curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

      curr_y = doc.y + (lineHeight*1.25)
    }

    // 91 = ExtBldg, 92 = EnviBldg 
    for(var i=91; i<= 92; i++){
      q_key = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 0}) ].v
      question = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 1}) ].v.replace(' (Answer)', '')
      answer = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': Row}) ].v

      doc.font(opts.font.weight.bold)
      doc.text(question, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

      curr_y = doc.y + (lineHeight*1.25)
      doc.font(opts.font.weight.normal)
      doc.text('Your Score:', curr_x, curr_y, {'align': 'left'})

      doc.font(opts.font.weight.bold)
      doc.text(answer, (curr_x + max_width), curr_y, {'align': 'left'})

      curr_y = doc.y
      doc.font(opts.font.weight.normal)
      doc.text('MLC/Garvan FP Average:', curr_x, curr_y, {'align': 'left'})
      doc.text(ScoreAve[q_key]['mlc'], (curr_x + max_width), curr_y, {'align': 'left'})

      curr_y = doc.y
      doc.text('Industry Average:', curr_x, curr_y, {'width': max_width, 'align': 'left'})
      doc.text(ScoreAve[q_key]['industry'], (curr_x + max_width), curr_y, {'align': 'left'})

      curr_y = doc.y + (lineHeight*1.25)
    }

    // Ee
    curr_y = doc.y + (lineHeight*2)

    doc.font(opts.font.weight.bold)
    doc.text('Overall Environment Score', curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

    lineHeight = doc.currentLineHeight()

    q_key = worksheet[ XLSX.utils.encode_cell({'c': 79, 'r': 0}) ].v.replace(' (Value)', '')
    answer = worksheet[ XLSX.utils.encode_cell({'c': 79, 'r': Row}) ].v
    doc.fontSize(opts.font.size.normal)

    max_width = doc.widthOfString('MLC/Garvan FP Environment Average:') + 5

    curr_y = doc.y + (lineHeight*1.5)
    doc.text('Your Environment Score:', curr_x, curr_y)
    doc.text(answer, (curr_x + (max_width - 40)), curr_y)

    doc.font(opts.font.weight.normal)
    
    curr_y = doc.y
    doc.text('MLC/Garvan FP Environment Average:', curr_x, curr_y)
    doc.text(ScoreAve[q_key]['mlc'], (curr_x + max_width), curr_y)

    curr_y = doc.y
    doc.text('Industry Environment Average:', curr_x, curr_y)
    doc.text(ScoreAve[q_key]['industry'], (curr_x + max_width), curr_y)


    // Follow up page
    doc.addPage(pageOptions)
    // footer and header
    doc.path(footer_svg_path).fillAndStroke("#c95109")
    doc.path(header_svg_path).fillAndStroke("#febf38")

    doc.lineWidth(20)
    doc.lineJoin('round')

    curr_x = opts.page.margin.left;
    curr_y = opts.page.margin.top;

    // standard opening page heading for the succedding pages
    box_width = opts.page.width-(opts.page.margin.left*2);
    box_height = opts.page.margin.top*0.2;
    doc.rect(curr_x, curr_y, box_width, box_height)
    doc.fillAndStroke('#febf38')
    doc.fillColor('#000')

    curr_x = opts.page.margin.left + 5;
    curr_y = opts.page.margin.top + 2;

    doc.font(opts.font.weight.bold)
    doc.fontSize(opts.font.size.normal)
    doc.text('Follow up', curr_x, curr_y, {'width': box_width, 'align': 'left'})

    doc.font(opts.font.weight.bold)
    doc.fontSize(opts.font.size.h1)

    doc.fontSize(opts.font.size.normal)
    lineHeight = doc.currentLineHeight()

    curr_y = doc.y + (lineHeight*2)

    // FollowUp = 94, 95 = DaysFollow, 96 = HowFollow
    for(var i=94; i<= 96; i++){
      q_key = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 0}) ].v
      question = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 1}) ].v.replace(' (Answer)', '')
      answer = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': Row}) ].v

      doc.font(opts.font.weight.bold)

      doc.text(question, curr_x, curr_y, {'width': box_width, 'align': 'left'})

      doc.font(opts.font.weight.normal)

      //get max width for the questions
      max_width = doc.widthOfString('MLC/Garvan FP Average: ') + 5

      // track current y-coord
      curr_y = doc.y + (lineHeight*1.5)
      doc.text('Your Score:', curr_x, curr_y, {'width': max_width, 'align': 'left'})

      doc.font(opts.font.weight.bold)

      doc.text(answer, (doc.x + max_width), curr_y, {'width': (box_width*0.66), 'align': 'left'})

      doc.font(opts.font.weight.normal)

      curr_y = doc.y
      doc.text('MLC/Garvan FP Average:', curr_x, curr_y, {'width': max_width, 'align': 'left'})
      doc.text(ScoreAve[q_key]['mlc'], (doc.x + max_width), curr_y, {'width': (box_width*0.66), 'align': 'left'})

      curr_y = doc.y
      doc.text('Other responses:', curr_x, curr_y, {'width': max_width, 'align': 'left'})
      doc.text(answer, (doc.x + max_width), curr_y, {'width': (box_width*0.66), 'align': 'left'})

      curr_y = doc.y
      doc.text('Industry Average:', curr_x, (curr_y), {'width': max_width, 'align': 'left'})
      doc.text(ScoreAve[q_key]['industry'], (doc.x + max_width), curr_y, {'width': (box_width*0.66), 'align': 'left'})

      curr_y = doc.y + (lineHeight*1.5)
    }

    // Interpreting These Results page
    doc.addPage(pageOptions)
    // footer and header
    doc.path(footer_svg_path).fillAndStroke("#c95109")
    doc.path(header_svg_path).fillAndStroke("#febf38")

    doc.lineWidth(20)
    doc.lineJoin('round')

    curr_x = opts.page.margin.left;
    curr_y = opts.page.margin.top;

    // standard opening page heading for the succedding pages
    box_width = opts.page.width-(opts.page.margin.left*2);
    box_height = opts.page.margin.top*0.2;
    doc.rect(curr_x, curr_y, box_width, box_height)
    doc.fillAndStroke('#febf38')
    doc.fillColor('#000')

    curr_x = opts.page.margin.left + 5;
    curr_y = opts.page.margin.top + 2;

    doc.font(opts.font.weight.bold)
    doc.fontSize(opts.font.size.h1)
    doc.text('Interpreting These Results', curr_x, curr_y, {'width': box_width, 'align': 'left'})

    doc.font(opts.font.weight.normal)
    doc.fontSize(opts.font.size.normal)

    lineHeight = doc.currentLineHeight()
    box_par_width = box_width - 10
    curr_y = doc.y + box_height + (lineHeight*1.5)

    doc.text("The feedback and ratings in this scorecard were provided by a potential client that contacted your business seeking financial advice. This person was recruited and screened to ensure they are a real potential client in the market for advice.", curr_x, curr_y, {'width': box_par_width, 'align': 'left'})
    doc.text("This person was renumerated with a flat payment but was not remunerated for any Statement of Advice, if they proceeded to engage the planner.", curr_x, (doc.y + lineHeight), {'width': box_par_width, 'align': 'left'})
    doc.text("The contents of this scorecard represent this person’s opinions and perceptions. Opinions and perceptions are not necessarily ‘facts’, however the overall impression that is left in the mind of this person will guide their future behaviour ‘regardless of the facts’. It is their impression that will guide their decision to deal with your business again or recommend your service.", curr_x, (doc.y + lineHeight), {'width': box_par_width, 'align': 'left'})
    doc.text("One person’s opinion (good or bad) needs to be considered in the context of the rest of your business performance, and is an opportunity to ‘reality check’ your customer acquisition process and identify any areas for improvement.", curr_x, (doc.y + lineHeight), {'width': box_par_width, 'align': 'left'})
    doc.text("In interpreting your result you should at least do the following:", curr_x, (doc.y + lineHeight), {'width': box_par_width, 'align': 'left'})

    var data = "1. Review your ‘Intention’ scores and the provided reasons (if relevant), as it is these scores that will tell you what overall impression you left with the client."
 
    doc.font(opts.font.weight.bold)
    doc.text(data.slice(0, 2), curr_x, (doc.y + lineHeight), {'width': box_par_width, 'align': 'left', 'continued': true})
    doc.font(opts.font.weight.normal)
    doc.text(data.slice(2, 14), {'continued': true})
    doc.font(opts.font.weight.bold)
    doc.text(data.slice(14, 33), {'continued': true})
    doc.font(opts.font.weight.normal)
    doc.text(data.slice(33))

    data = "2. Consider your high and low scores in the other sections, to build picture of where you can improve. Compare your performance to MLC/Garvan FP average so you are aware of where you stand relative to your peers."

    doc.font(opts.font.weight.bold)
    doc.text(data.slice(0, 2), curr_x, (doc.y + lineHeight), {'width': box_par_width, 'align': 'left', 'continued': true})
    doc.font(opts.font.weight.normal)
    doc.text(data.slice(2, 103), {'continued': true})
    doc.font(opts.font.weight.bold)
    doc.text(data.slice(103, 127), {'continued': true})
    doc.font(opts.font.weight.normal)
    doc.text(data.slice(127))

    data = "Questions and Feedback about the format of this report or the process used to conduct this mystery shopping report can be directed to MLC/Garvan FP, who will liaise with CoreData if necessary.";

    doc.font(opts.font.weight.bold)
    doc.text(data.slice(0, 23), curr_x, (doc.y + lineHeight), {'width': box_par_width, 'align': 'left', 'continued': true})
    doc.font(opts.font.weight.normal)
    doc.text(data.slice(23))

    // final page
    doc.addPage(pageOptions)

    // footer and header
    doc.path('M 0 '+ (opts.page.height) +' l 0 -' + (footer_band_width+20) + ' l ' + opts.page.width + ' 0 l 0 ' + (footer_band_width + 20) + ' l -'+ opts.page.width +' 0').fillAndStroke("#5a5a5a")
    doc.path(header_svg_path).fillAndStroke("#febf38")

    //doc.pipe(res)

    // Finalize PDF file
    doc.end()

  }

  // force the requesting resource to download the pdf
  //res.setHeader('Content-disposition', 'attachment; filename=cer-report.pdf');


  res.send('processed ranges:' + JSON.stringify(range));
});

app.listen(8000, function () {
  console.log('Example app listening on port 8000!');
});
