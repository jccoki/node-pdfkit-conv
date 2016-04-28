/* vim: sw=2 ts=2 expandtab */
var fs = require('fs');

var express = require('express');
var app = express();

var XLSX = require('xlsx');
var PDFDocument = require('pdfkit');

var norm_func = function(num){
//  return Math.round(num)
  return num
}

var repl_func = function(str){
  str = str.replace(' (Answer)', '')
  str = str.replace(' (Text)', '')
  str = str.replace(' (Value)', '')

  return str
}

var lines_func = function(text_width, writable_width){
  // add more line to accommodate running paragraph text
  var lines = Math.ceil(text_width/writable_width)
//  if(lines > 4){
//    lines = lines + 1.5
//  }

  return lines
}

function getMaxOfArray(numArray) {
  return Math.max.apply(null, numArray);
}

var command_args = process.argv
var start_row = command_args[2]
//var start_row = 306
var excel_file = 'WW\ 2016\ Overall\ Individual\ Adviser\ Report\ Data\ v3.1.xlsx'
var input_sheet = 'Individual Report';

var target_output_dir = './output'
var licensee_logo = 'count_logo.png'
var licensee_images = 'resources/WW\ Licensees\ Charts\ v1.4/Count Financial'
var licensee_front_page = 'sunrise-274257_1280.jpg'

// report theme
var header_fill_style= '#255fa5'
var footer_fill_style= '#5a5a5a'
var table_header_color = '#fff'

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
      'bottom': 14.4,
      'left': 90,
      'right': 77.76
    },
    'padding': {
      'top': 32,
      'bottom': 2.88,
      'left': 32,
      'right': 32,
    }
  },
  'font': {
    'weight': {
      'bold': 'node_modules/pdfkit/js/font/data/Arial_Bold.ttf',
      'italic': 'node_modules/pdfkit/js/font/data/Arial_Italic.ttf',
      'boldItalic': 'node_modules/pdfkit/js/font/data/Arial_Bold_Italic.ttf',
      'normal': 'node_modules/pdfkit/js/font/data/Arial.ttf'},
    'size': {
      'normal': 12,
      'h1': 14,
      'footer': 6}
    }
};

var bulleted_items

var licensee_full_text = ''
var licensee_name = ''

//app.get('/', function (req, res) {
  // start reading the XLSX file
  var workbook = XLSX.readFile(excel_file);
  //excel file may contain multiple sheets
  var sheet_name_list = workbook.SheetNames;
  // get the range for rows and columns
  var ranges = workbook.Sheets[input_sheet]['!ref'];
  var range = XLSX.utils.decode_range(ranges);
  var worksheet = workbook.Sheets[input_sheet]

  // save the PDF properties variables
  var header_band_height = 20;
  var footer_band_height = 20;
  var pageOptions = {
    'size': [opts.page.width, opts.page.height],
    'margins': {'top': 0, 'left': 0, 'right': 0, 'bottom': 0},
    'bufferPages': true
  };

  // parsed at column 6
  var plannerName = '';

  // allow space between the table border and the bounding rect
  // -10pts for left margin, -7.5pts for right margin
  // set the table heading height
  var tableOpts = {
    'dividerWidth': 2,
    'margin': opts.page.margin.left + 15
  }

  var c_addr = '',
      question = '',
      answer = '',
      q_key = '',
      data = ''

  var pageRange = 0

  // variables for bounding rect box
  var box_height, box_width, box_par_width
  var licensee_rate

  // read the row data one by one
  // start of usable row is 2, 0-index rule
  // retrieving empty contents from cell address returns undefined
//  for(var Row = 2; Row <= range.e.r; Row++) {
  var start_row = start_row - 1
  var end_row = start_row
  // questions are in 2nd row, answers are in current [row, column]
  for(var Row = start_row; Row <= end_row; Row++) {
    number = worksheet[ XLSX.utils.encode_cell({'c': 0, 'r': Row}) ].v
    // limit to 4 digit filename
    // number = ("0000" + number).substr(-4,4)
    output_dir = number

    // fetch licensee full text
    licensee_full_text = worksheet[ XLSX.utils.encode_cell({'c': 5, 'r': Row}) ].v
    // fetch licensee full name
    licensee_name = worksheet[ XLSX.utils.encode_cell({'c': 98, 'r': Row}) ].v
    // fetch planner name
    plannerName = worksheet[ XLSX.utils.encode_cell({'c': 6, 'r': Row}) ].v

    var views_planner_offer_questions = []

    // mystery shopper profile questions
    mystery_shopper_profile_questions = [];

    // 8 = UsageFA, 9 = LastTimeFA, 10 = PayMethod
    for(var Col = 8; Col <= 10; Col++){
      c_addr = XLSX.utils.encode_cell({'c': Col, 'r': Row});
//      question = worksheet[ XLSX.utils.encode_cell({'c': Col, 'r': 1}) ].v.replace(' (Answer)', '')
      switch(Col){
        case 8:
          question = 'What best describes your usage of financial planners (excluding basic accountancy)?'
          break;
        case 9:
          question = 'When was the last time you used a financial planner? (excluding this project)'
          break;
        case 10:
          question = 'Do you have a preference for the way in which you pay for advice? (e.g. commission, flat fee, fee based on a percentage of assets under management, etc.)'
          break;
      }

      var concat
      // column 10 and 11 should be concatenated
      // intialize to empty
      answer = ''
      if(worksheet[c_addr]){
        answer = worksheet[c_addr].v
        if(Col == 10){
          if(worksheet[XLSX.utils.encode_cell({'c': (Col + 1), 'r': Row})]){
            answer = answer + worksheet[XLSX.utils.encode_cell({'c': (Col + 1), 'r': Row})].v
          }
        }
      }else{
        answer = 'Not applicable'
      }
      mystery_shopper_profile_questions.push({'question': question,'answer': answer})
    }

    // start of column = 12, end of column = 21, row = 1
    for(var Col = 12; Col <= 21; Col++){
      c_addr = XLSX.utils.encode_cell({'c': Col, 'r': Row});

      // retrieve question key
      q_key = worksheet[ XLSX.utils.encode_cell({'c': Col, 'r': 0}) ].v

      // retrieve question
      question = repl_func( worksheet[ XLSX.utils.encode_cell({'c': Col, 'r': 1}) ].v )
      answer = worksheet[c_addr].v

      // @TODO refactor and make this better
      switch(q_key){
        case 'HiRet':
          licensee_rate = worksheet[ XLSX.utils.encode_cell({'c': 103, 'r': Row}) ].v
          industry_rate = '58%'
          break;
        case 'AccMrkt':
          licensee_rate = worksheet[ XLSX.utils.encode_cell({'c': 104, 'r': Row}) ].v
          industry_rate = '82%'
          break;
        case 'ComplMrkt':
          licensee_rate = worksheet[ XLSX.utils.encode_cell({'c': 105, 'r': Row}) ].v
          industry_rate = '55%'
          break;
        case 'RsrchMrkt':
          licensee_rate = worksheet[ XLSX.utils.encode_cell({'c': 106, 'r': Row}) ].v
          industry_rate = '52%'
          break;
        case 'AccFunds':
          licensee_rate = worksheet[ XLSX.utils.encode_cell({'c': 107, 'r': Row}) ].v
          industry_rate = '52%'
          break;
        case 'HelpMnge':
          licensee_rate = worksheet[ XLSX.utils.encode_cell({'c': 108, 'r': Row}) ].v
          industry_rate = '62%'
          break;
        case 'DevStrat':
          licensee_rate = worksheet[ XLSX.utils.encode_cell({'c': 109, 'r': Row}) ].v
          industry_rate = '70%'
          break;
        case 'DcsnMkg':
          licensee_rate = worksheet[ XLSX.utils.encode_cell({'c': 110, 'r': Row}) ].v
          industry_rate = '44%'
          break;
        case 'SaveTime':
          licensee_rate = worksheet[ XLSX.utils.encode_cell({'c': 111, 'r': Row}) ].v
          industry_rate = '45%'
          break;
        case 'ReAssure':
          licensee_rate = worksheet[ XLSX.utils.encode_cell({'c': 112, 'r': Row}) ].v
          industry_rate = '39%'
          break;
      }

      views_planner_offer_questions.push({
        'qkey': q_key,
        'question': question,
        'answer': answer ,
        'licensee': licensee_rate + '%',
        'industry': industry_rate
      })
    }

    if( worksheet[ XLSX.utils.encode_cell({'c': 22, 'r': Row}) ] ){
      answer = worksheet[ XLSX.utils.encode_cell({'c': 22, 'r': Row}) ].v
    }else{
      answer = ''
    }

    var BenefFA = {
      'question': repl_func( worksheet[ XLSX.utils.encode_cell({'c': 22, 'r': 1}) ].v ),
      'answer': answer}

    var TopExpectations = []
    // start of column = 23, end of column = 25, row = 1
    for(var Col = 23; Col <= 25; Col++){
      key = worksheet[ XLSX.utils.encode_cell({'c': Col, 'r': 0}) ].v
      switch(key){
        case 'Top1':
          industry_top_expect = 'Honesty'
          licensee_ave = worksheet[ XLSX.utils.encode_cell({'c': 100, 'r': Row}) ].v
          break;
        case 'Top2':
          industry_top_expect = 'Value for money'
          licensee_ave = worksheet[ XLSX.utils.encode_cell({'c': 101, 'r': Row}) ].v
          break;
        case 'Top3':
          industry_top_expect = 'Maximise investment returns'
          licensee_ave = worksheet[ XLSX.utils.encode_cell({'c': 102, 'r': Row}) ].v
          break;
      }

      TopExpectations.push({
        'key': key,
        'question': worksheet[ XLSX.utils.encode_cell({'c': Col, 'r': 1}) ].v,
        'answer': worksheet[ XLSX.utils.encode_cell({'c': Col, 'r': Row}) ].v,
        'licensee': licensee_ave,
        'industry': industry_top_expect
      })
    }

    var parOpts

    // start generating pdf
    var doc = new PDFDocument(pageOptions)

    var totalWidth = doc.page.width-(opts.page.margin.left*2)
    var header_svg_path = 'M 0 0 l 0 ' + header_band_height + ' l ' + doc.page.width + ' 0 l 0 -' + header_band_height + ' l -'+ doc.page.width +' 0';

    var footer_svg_path = 'M 0 '+ doc.page.height +' l 0 -' + footer_band_height + ' l ' + doc.page.width + ' 0 l 0 ' + footer_band_height + ' l -'+ doc.page.width +' 0';

    // create header and footer everytime a new page is added
    doc.on('pageAdded', function(){
      // footer and header
      doc.path(footer_svg_path).fillAndStroke(footer_fill_style)
      doc.path(header_svg_path).fillAndStroke(header_fill_style)

      doc.fontSize(10)
      doc.font(opts.font.weight.normal)

      curr_y = (opts.page.height-(footer_band_height*2))

      pageNum = (doc.bufferedPageRange().start + doc.bufferedPageRange().count) - 1

      doc.fillColor('#000')
      curr_x =  doc.page.width - (opts.page.margin.right + 10)
      doc.text(pageNum, curr_x, curr_y)

      curr_x = opts.page.margin.left
      doc.text(plannerName, curr_x, curr_y)
    })

    // header
    doc.path(header_svg_path).fillAndStroke(header_fill_style)
    doc.image('resources/coredata-logo.png', (opts.page.width*0.65), (header_band_height + (header_band_height*2)), {'scale': 0.09})
    // position the MLC logo 18% of the total page height and within 25% around the middle of the page
// doc.image('resources/ww-mlc-logo.png', ((opts.page.width/2)-(opts.page.width*.125)), (opts.page.height*0.18), {'scale': 0.35})

    // draw the middle band line starting from 33% of the total page height
//    doc.path('M 0 '+ (opts.page.height*0.33) +' l 0 -' + header_band_height + ' l ' + opts.page.width + ' 0 l 0 ' + header_band_height + ' l -'+ opts.page.width +' 0').fillAndStroke("#c95109");

    // front page logos
    doc.image('resources/logos/' + licensee_logo, (opts.page.width*0.375), (opts.page.height*0.20), {'scale': 0.55})
//    doc.image('resources/ww-frontpage-bg.png', (opts.page.width*0.175), (opts.page.height*0.25), {'scale': 0.45})
    doc.image('resources/front pages/' + licensee_front_page, 0, (opts.page.height*0.27), {'width': doc.page.width, 'height': opts.page.height*0.325})

    // yellow foreground
    doc.rect(0, ((opts.page.height*0.60)-10), opts.page.width, (opts.page.height*0.40)+10).fillAndStroke(header_fill_style)
    // draw the page title and its bounding rect box
    // black background
    doc.path('M 5 '+ (opts.page.height-5) + ' l 0 -' + (opts.page.height*0.40) + ' l ' + (opts.page.width-10) + ' 0 l 0 ' + (opts.page.height*0.40) + ' l -'+ (opts.page.width-10) +' 0')
    doc.fillAndStroke("#272727")

    doc.fontSize(22)
    doc.font(opts.font.weight.bold)
    doc.fillColor(header_fill_style)

    // set the position of cursor for text alignment
    curr_x = opts.page.margin.left*1.25
    curr_y = opts.page.height*0.75
    doc.text('Financial Planner Shadow Shopping', curr_x, curr_y)

    lineHeight = doc.currentLineHeight()
    curr_x = doc.x
    curr_y = doc.y + lineHeight
    
    doc.fontSize(16)
    doc.font(opts.font.weight.normal)
    doc.fillColor('#fff')
    doc.text('Individual Planner Scorecard: ' + plannerName, curr_x, curr_y)

    // INTRODUCTION page
    doc.addPage(pageOptions)

    curr_x = opts.page.margin.left;
    curr_y = opts.page.margin.top;

    box_par_width = opts.page.width - (opts.page.margin.left + opts.page.margin.right)
    parOpts = {'width': box_par_width, 'align': 'left'}

    doc.fontSize(20)
    doc.font(opts.font.weight.bold)
    doc.fillColor('#000')
    doc.text('INTRODUCTION', curr_x, curr_y, parOpts)

    doc.fontSize(opts.font.size.normal)
    doc.font(opts.font.weight.normal)
    lineHeight = doc.currentLineHeight()

    curr_y = doc.y + lineHeight
    data = 'CoreData\’s annual financial planning shadow shopping study of the Australian advice industry has been running since 2003 and is designed to give licensees an accurate snapshot of how well their distribution channels are performing in both absolute and relative terms.'
    doc.text(data, curr_x, curr_y, parOpts)

    curr_y = doc.y + lineHeight
    data = 'This year\’s Industry study incorporated 235 shops of planners representing 13 leading Australian licensees.'
    doc.text(data, curr_x, curr_y, parOpts)

    curr_y = doc.y + lineHeight
    data = 'CoreData recruits shoppers who:'
    doc.text(data, curr_x, curr_y, parOpts)

    //bulleted items
    bulleted_items = []
    bulleted_items[0] = 'are looking for financial advice'
    bulleted_items[1] = 'do not have a financial planner or are open to moving from their current planner'
    bulleted_items[2] = 'are 35 years and above'
    bulleted_items[3] = 'have minimum investments of $150,000 (including superannuation)'

    for(var i=0; i<bulleted_items.length; i++ ){
      curr_x = opts.page.margin.left + 9
      curr_y = doc.y
      doc.circle(curr_x, (curr_y + (lineHeight/2.5)), 2).fill('#000').stroke()
      doc.text(bulleted_items[i], (curr_x + 9), curr_y, {'width': box_par_width - 18, 'align': 'left'})
    }

    curr_x = opts.page.margin.left;
    curr_y = doc.y + lineHeight
    data = 'The shoppers go through the process of making an appointment, meeting with a planner, discussing their situation, outlining next steps, assessing the planner\’s abilities and deciding whether to progress to the next stage.'
    doc.text(data, curr_x, curr_y, parOpts)

    curr_y = doc.y + lineHeight
    data = 'Shoppers are remunerated with a flat payment whether or not they subsequently engage the planner or seek a statement of advice.'
    doc.text(data, curr_x, curr_y, parOpts)

    curr_y = doc.y + lineHeight
    data = 'Shopper responses are summarised using CoreData\’s ACQUIRE\© Index which measures performance across seven areas, namely:'
    doc.text(data, curr_x, curr_y, parOpts)

    bulleted_items = []
    bulleted_items[0] = 'Assurances'
    bulleted_items[1] = 'Compliance'
    bulleted_items[2] = 'Quality'
    bulleted_items[3] = 'Understanding'
    bulleted_items[4] = 'Intention'
    bulleted_items[5] = 'Reaction'
    bulleted_items[6] = 'Environment'

    for(var i=0; i<bulleted_items.length; i++ ){
      curr_x = opts.page.margin.left + 9
      curr_y = doc.y
      doc.circle(curr_x, (curr_y + (lineHeight/2.5)), 2).fill('#000').stroke()
      doc.text(bulleted_items[i], (curr_x + 9), curr_y, {'width': box_par_width - 18, 'align': 'left'})
    }

    // OVERALL RESULTS page
    doc.addPage(pageOptions)

    curr_x = opts.page.margin.left;
    curr_y = opts.page.margin.top;

    box_par_width = opts.page.width - (opts.page.margin.left + opts.page.margin.right)
    parOpts = {'width': box_par_width, 'align': 'left'}

    doc.fontSize(20)
    doc.font(opts.font.weight.bold)
    doc.fillColor('#000')

    doc.text('OVERALL RESULTS', curr_x, curr_y, parOpts)

    doc.fontSize(opts.font.size.normal)
    doc.font(opts.font.weight.normal)
    lineHeight = doc.currentLineHeight()

    curr_y = doc.y + lineHeight

    doc.text('The overall ACQUIRE scores for you, your licensee and the Industry are set out below:', curr_x, curr_y, parOpts)

    curr_x = opts.page.margin.left
    curr_y = doc.y + lineHeight
    box_height = lineHeight*2

    table_heading = ['Your score', licensee_name, 'Industry Average']
    // this should load the data
    table_row = []
    table_row[0] = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 7, 'r': Row}) ].v )
    table_row[1] = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 99, 'r': Row}) ].v )
    table_row[2] = '76'

    doc.font(opts.font.weight.bold)
    doc.lineWidth(0.3)
    for(var i = 0; i<table_heading.length; i++){
      doc.rect(curr_x, curr_y, (box_par_width/3), box_height)
      doc.fillAndStroke(header_fill_style, '#000')
      doc.fillColor(table_header_color)
      doc.text(table_heading[i], curr_x, (curr_y+(box_height*0.30)), {'width': (box_par_width/3), 'align': 'center'})
      curr_x = curr_x + (box_par_width/3)
    }

    // position relative to the table header
    curr_y = curr_y + box_height

    doc.fontSize(36)
    lineHeight = doc.currentLineHeight()
    box_height = lineHeight*2

    curr_x = opts.page.margin.left
    for(var i = 0; i<table_row.length; i++){
      doc.rect(curr_x, curr_y, (box_par_width/3), box_height)
      doc.fillAndStroke('#fff', '#000')
      doc.fillColor('#000')
      doc.text(table_row[i], curr_x, (curr_y+(lineHeight*0.50)), {'width': (box_par_width/3), 'align': 'center'})

      curr_x = curr_x + (box_par_width/3)
    }

    doc.fontSize(opts.font.size.normal)
    doc.font(opts.font.weight.normal)
    lineHeight = doc.currentLineHeight()

    curr_x = opts.page.margin.left

    curr_y = doc.y + (lineHeight*2)
    data = 'This year\’s study again reinforced the notion that there is no single driver of perceived value and the value of advice.'
    doc.text(data, curr_x, curr_y, parOpts)

    curr_y = doc.y + lineHeight
    data = 'Numerous factors interact to deliver the whole experience and overall perceptions of value. Together these factors deliver a result much greater than just the sum of the constituent parts. This is the Gestalt Phenomenon.'
    doc.text(data, curr_x, curr_y, parOpts)

    curr_y = doc.y + lineHeight
    data = 'The challenge is that this works both ways. Weakness in any one area will significantly diminish the overall value created. A focus on just one or two factors is unlikely to lead to an increase in perceived value.'
    doc.text(data, curr_x, curr_y, parOpts)

    curr_y = doc.y + lineHeight
    data = 'That said, some factors are more related and dependent on each other and some are more important. Our research identified six factors that in combination determine more than 70% of the perceived value of advice, namely:'
    doc.text(data, curr_x, curr_y, parOpts)

    var str_index

    data = 'Advice was suited to my needs'
    str_index = data.indexOf('my')
    curr_x = opts.page.margin.left + 9
    curr_y = doc.y
    doc.circle(curr_x, (curr_y + (lineHeight/2.5)), 2).fill('#000').stroke()
    doc.text(data.slice(0, str_index), (curr_x + 9), curr_y, {'width': box_par_width - 18, 'align': 'left', 'continued': true})
    doc.text(data.slice(str_index, str_index + 2), (curr_x + 9), curr_y, {'width': box_par_width - 18, 'align': 'left', 'continued': true, 'underline': true})
    doc.text(data.slice(str_index + 2, data.length), (curr_x + 9), curr_y, {'width': box_par_width - 18, 'align': 'left', 'underline': false})

    data = 'Planner demonstrated keenness for my business'
    str_index = data.indexOf('my')
    curr_x = opts.page.margin.left + 9
    curr_y = doc.y
    doc.circle(curr_x, (curr_y + (lineHeight/2.5)), 2).fill('#000').stroke()
    doc.text(data.slice(0, str_index), (curr_x + 9), curr_y, {'width': box_par_width - 18, 'align': 'left', 'continued': true})
    doc.text(data.slice(str_index, str_index + 2), (curr_x + 9), curr_y, {'width': box_par_width - 18, 'align': 'left', 'continued': true, 'underline': true})
    doc.text(data.slice(str_index + 2, data.length), (curr_x + 9), curr_y, {'width': box_par_width - 18, 'align': 'left', 'underline': false})

    data = 'Planner was trustworthy'
    curr_x = opts.page.margin.left + 9
    curr_y = doc.y
    doc.circle(curr_x, (curr_y + (lineHeight/2.5)), 2).fill('#000').stroke()
    doc.text(data, (curr_x + 9), curr_y, {'width': box_par_width - 18, 'align': 'left'})

    data = 'Demonstrated ability to influence me'
    str_index = data.indexOf('me')
    curr_x = opts.page.margin.left + 9
    curr_y = doc.y
    doc.circle(curr_x, (curr_y + (lineHeight/2.5)), 2).fill('#000').stroke()
    doc.text(data.slice(0, str_index), (curr_x + 9), curr_y, {'width': box_par_width - 18, 'align': 'left', 'continued': true})
    doc.text(data.slice(str_index, str_index + 2), (curr_x + 9), curr_y, {'width': box_par_width - 18, 'align': 'left', 'continued': true, 'underline': true})
    doc.text(data.slice(str_index + 2, data.length), (curr_x + 9), curr_y, {'width': box_par_width - 18, 'align': 'left', 'underline': false})

    data = 'Comprehensive knowledge of advice strategies and investments'
    curr_x = opts.page.margin.left + 9
    curr_y = doc.y + lineHeight
    doc.circle(curr_x, (curr_y + (lineHeight/2.5)), 2).fill('#000').stroke()
    doc.text(data, (curr_x + 9), curr_y, {'width': box_par_width - 18, 'align': 'left'})

    data = 'Demonstrated effectively the benefits of the services and advice to me'
    str_index = data.indexOf('me')
    curr_x = opts.page.margin.left + 9
    curr_y = doc.y
    doc.circle(curr_x, (curr_y + (lineHeight/2.5)), 2).fill('#000').stroke()
    doc.text(data.slice(0, str_index), (curr_x + 9), curr_y, {'width': box_par_width - 18, 'align': 'left', 'continued': true})
    doc.text(data.slice(str_index, str_index + 2), (curr_x + 9), curr_y, {'width': box_par_width - 18, 'align': 'left', 'continued': true, 'underline': true})
    doc.text(data.slice(str_index + 2, data.length), (curr_x + 9), curr_y, {'width': box_par_width - 18, 'align': 'left', 'underline': false})

    curr_x = opts.page.margin.left
    curr_y = doc.y + (lineHeight*2)
    data = 'What is clear is that advisers who perform well in all of these areas are perceived to create more value. Critical to this is how advisers help the individual client to address their specific needs, goals and objectives.'
    doc.text(data, curr_x, curr_y, parOpts)

    curr_y = doc.y + lineHeight
    data = 'Advisers should reflect on how they can construct the first meeting to ensure these factors are explicitly addressed. They should talk specifically about these things: \“This advice will address your needs in the following ways...\”, \“I am keen to help you achieve these outcomes\”, etc.'
    doc.text(data, curr_x, curr_y, parOpts)

    curr_y = doc.y + lineHeight
    data = 'A sustained focus on this basket of factors will lead to enhanced perceptions of value.'
    doc.text(data, curr_x, curr_y, parOpts)

   //doc.flushPages()

    // Interpreting These Results page
    doc.addPage(pageOptions)

    curr_x = opts.page.margin.left;
    curr_y = opts.page.margin.top;

    box_par_width = opts.page.width - (opts.page.margin.left + opts.page.margin.right)
    parOpts = {'width': box_par_width, 'align': 'left'}

    doc.fontSize(20)
    doc.font(opts.font.weight.bold)
    doc.text('INTERPRETING THESE RESULTS', curr_x, curr_y, parOpts)

    doc.fontSize(opts.font.size.normal)
    doc.font(opts.font.weight.normal)
    lineHeight = doc.currentLineHeight()

    curr_y = doc.y + lineHeight

    doc.text('This scorecard sets out one person\’s opinions and perceptions and not necessarily \‘facts\’.  They reflect the overall impressions left in the mind of this person that will guide their decision to deal with your business again or recommend your service to others.', curr_x, curr_y, parOpts)

    curr_x = opts.page.margin.left;
    curr_y = doc.y + lineHeight
    doc.text('One person\’s opinion (good or bad) needs to be considered in the context of the rest of your business performance, and is an opportunity to \‘reality check\’ your customer acquisition process and identify any areas for improvement.', curr_x, curr_y, parOpts)

    curr_x = opts.page.margin.left;
    curr_y = doc.y + lineHeight
    doc.text("Bullet charts should be interpreted as follows:", curr_x, curr_y, parOpts)

    curr_y = doc.y
    doc.image('resources/bullet_chart_howto.png', curr_x, curr_y, {'scale': 0.575})

    curr_y = doc.y + lineHeight

    doc.text("In interpreting your results you should at least do the following:", curr_x, curr_y, parOpts)

    var data = "1. Review your ‘Intention’ scores and the provided reasons (if relevant), as it is these scores that will tell you what overall impression you left with the client."
 
    doc.font(opts.font.weight.bold)
    doc.text(data.slice(0, 2), curr_x, (doc.y + lineHeight), {'width': box_par_width, 'align': 'left', 'continued': true})
    doc.font(opts.font.weight.normal)
    doc.text(data.slice(2, 14), {'continued': true})
    doc.font(opts.font.weight.bold)
    doc.text(data.slice(14, 33), {'continued': true})
    doc.font(opts.font.weight.normal)
    doc.text(data.slice(33))

    // @TODO this is not correct, licensee average data is still missing
    licensee_ave = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 7, 'r': Row}) ].v )
    data = "2. Consider your high and low scores in the other sections, to build picture of where you can improve. Compare your performance to "+ licensee_full_text +" average so you are aware of where you stand relative to your peers."

    doc.font(opts.font.weight.bold)
    doc.text(data.slice(0, 2), curr_x, (doc.y + lineHeight), {'width': box_par_width, 'align': 'left', 'continued': true})
    doc.font(opts.font.weight.normal)
    doc.text(data.slice(2, 103), {'continued': true})
    doc.font(opts.font.weight.bold)
    doc.text(data.slice(103, 127), {'continued': true})
    doc.font(opts.font.weight.normal)
    doc.text(data.slice(127))

    data = "Questions and Feedback about the contents of this report or the process used to conduct this mystery shopping report can be directed to " + licensee_full_text + '.';

    curr_y = doc.y + lineHeight
    doc.font(opts.font.weight.bold)
    doc.text(data.slice(0, 23), curr_x, curr_y, {'width': box_par_width, 'align': 'left', 'continued': true})
    doc.font(opts.font.weight.normal)
    doc.text(data.slice(23))

   //doc.flushPages()


    // SHOPPER PROFILE page 1
    doc.addPage(pageOptions)

    // footer and header
    curr_x = opts.page.margin.left;
    curr_y = opts.page.margin.top;

    box_par_width = opts.page.width - (opts.page.margin.left + opts.page.margin.right)
    parOpts = {'width': box_par_width, 'align': 'left'}

    doc.fontSize(20)
    doc.font(opts.font.weight.bold)
    doc.fillColor('#000')
    doc.text('SHOPPER PROFILE', curr_x, curr_y, parOpts)

    doc.fontSize(opts.font.size.normal)
    doc.font(opts.font.weight.normal)
    lineHeight = doc.currentLineHeight()

    curr_y = doc.y + lineHeight

    doc.text('Prior to meeting with you, the shopper was asked a number of questions regarding their attitudes towards advice, including:', curr_x, curr_y, parOpts)

    curr_x = opts.page.margin.left + 3;
    curr_y = doc.y + lineHeight

    // iterate mystery_shopper_profile_questions
    doc.fontSize(opts.font.size.normal)
    for(var i = 0; i < mystery_shopper_profile_questions.length; i++){

      doc.font(opts.font.weight.boldItalic)
      doc.text(mystery_shopper_profile_questions[i].question, curr_x, curr_y, parOpts);

      curr_y = doc.y + lineHeight;

      doc.font(opts.font.weight.normal)
      doc.text('Answer:', curr_x, curr_y)

      doc.font(opts.font.weight.bold)
      doc.text(mystery_shopper_profile_questions[i].answer, (curr_x+(box_par_width*0.15)), curr_y, {'width': (box_par_width*0.8), 'align': 'left'})

      curr_y = doc.y + (lineHeight*2);
    }

    doc.font(opts.font.weight.boldItalic)
    doc.text('What are the top three expectations you would have from a financial planner you are considering using (i.e. what do they have to do to get your business)?', curr_x, curr_y, parOpts);

    thead_heading = [];
    thead_heading[0] = {'width': 0.14, 'data': ''};
    thead_heading[1] = {'width': 0.30, 'data': 'Your shopper'};
    thead_heading[2] = {'width': 0.28, 'data': licensee_name};
    thead_heading[3] = {'width': 0.28, 'data': 'Industry'};

    trow_heading = [];

    // assure correct sorting
    trow_heading[0] = 'Top 1';
    trow_heading[1] = 'Top 2';
    trow_heading[2] = 'Top 3';

    // track table coords
    doc.font(opts.font.weight.bold)
    lineHeight = doc.currentLineHeight()

    box_height = lineHeight*2
    curr_x = opts.page.margin.left + 5
    curr_y = doc.y + box_height

    // draw table heading
    doc.lineWidth(0.3)
    padding = lineHeight*0.45
    for(var i=0; i<thead_heading.length; i++){
       doc.rect(curr_x, curr_y, (totalWidth*thead_heading[i].width), box_height)
       doc.fillAndStroke(header_fill_style, '#000')
       doc.fillColor(table_header_color)
       doc.text(thead_heading[i].data, (curr_x + 2.5), (curr_y + padding), {'align': 'center', 'width': (totalWidth*thead_heading[i].width) - 5})

       curr_x += (totalWidth*thead_heading[i].width)
    }
    doc.font(opts.font.weight.normal)

    // draw table rows
    curr_y = curr_y + box_height

    // row level
    for(i=0; i<TopExpectations.length; i++){
      var max_line
      // dynamically allocation of table cell height
      writable_width = (totalWidth*thead_heading[1].width) - 20
      doc_width = doc.widthOfString(TopExpectations[i].answer)
      answer_lines = lines_func(doc_width, writable_width)
      // very special caser
      if(TopExpectations[i].answer == 'Keeps me up to date with changes in the investment and regulatory'){
        answer_lines = answer_lines + 2
      }

      if(TopExpectations[i].answer == 'Demonstrate knowledge and expertise'){
        answer_lines = answer_lines + 1
      }

      if(TopExpectations[i].answer == 'Takes time to listen and explain things to me'){
        answer_lines = answer_lines
      }

      if(TopExpectations[i].answer == 'Put my interests first'){
        answer_lines = answer_lines - 1
      }

      if(TopExpectations[i].answer == 'Transparency around fees and charges'){
        answer_lines = answer_lines - 1
      }

      if(TopExpectations[i].answer == 'The adviser\'s business is independent of the bigger institutions and/or banks'){
        answer_lines = answer_lines + 1
      }

      writable_width = (totalWidth*thead_heading[2].width) - 5
      doc_width = doc.widthOfString(TopExpectations[i].licensee)
      licensee_lines = lines_func(doc_width, writable_width)

      writable_width = (totalWidth*thead_heading[3].width) - 5
      doc_width = doc.widthOfString(TopExpectations[i].industry)
      industry_lines = lines_func(doc_width, writable_width)

      max_line = Math.max(answer_lines, licensee_lines, industry_lines) + 1
      box_height = lineHeight*max_line

      curr_x = opts.page.margin.left + 5
      // column level
      for(var j=0; j<thead_heading.length; j++){
        var style = (i % 2) == 1 ? '#fafafa' : '#dedede';
        doc.rect(curr_x, curr_y, (totalWidth*thead_heading[j].width), box_height)
        doc.fillAndStroke(style, '#000')

        curr_x += (totalWidth*thead_heading[j].width)
      }

      // reset x-coords before drawing text
      curr_x = opts.page.margin.left + 5

      q_key = TopExpectations[i].key

      doc.fillColor('#000')
      // draw the questions
      padding = ((max_line - 1)*lineHeight) * 0.5
      doc.font(opts.font.weight.bold)
      doc.text(TopExpectations[i].question,( curr_x + 2.5), (curr_y + padding), {'width': (totalWidth*thead_heading[0].width) - 5, 'align': 'center'});
      doc.font(opts.font.weight.normal)

      // draw the self rating
      // very special caser
      switch(TopExpectations[i].answer.toLowerCase()){
        case 'the adviser\'s business is independent of the bigger institutions and/or banks':
        case 'keeps me up to date with changes in the investment and regulatory environment':
        case 'demonstrate knowledge and expertise':
        case 'Takes time to listen and explain things to me':
          padding = lineHeight * 0.25
          break;
        default:
          padding = ((max_line*lineHeight) - (answer_lines*lineHeight)) * 0.5
          break;
      }

      curr_x = curr_x + (totalWidth*thead_heading[0].width)
      doc.text(TopExpectations[i].answer, (curr_x + 2.5), (curr_y + padding), {'width': (totalWidth*thead_heading[1].width) - 5, 'align': 'center'});

      // draw licensee average
      padding = ((max_line*lineHeight) - (licensee_lines*lineHeight)) * 0.5
      curr_x = curr_x+(totalWidth*thead_heading[1].width)
      doc.text(TopExpectations[i].licensee, (curr_x + 2.5), (curr_y + padding), {'width': (totalWidth*thead_heading[2].width) - 5, 'align': 'center'});

      // draw industry average
      padding = ((max_line*lineHeight) - (industry_lines*lineHeight)) * 0.5
      curr_x = curr_x+(totalWidth*thead_heading[2].width)
      doc.text(TopExpectations[i].industry, (curr_x + 2.5), (curr_y + padding), {'width': (totalWidth*thead_heading[3].width) - 5, 'align': 'center'});

      // reset x-coord to the leftmost writable page of the page
      curr_x = opts.page.margin.left
      curr_y = curr_y + box_height
    }


    // SHOPPER PROFILE page 2
    doc.addPage(pageOptions)

    box_par_width = opts.page.width - (opts.page.margin.left + opts.page.margin.right)
    parOpts = {'width': box_par_width, 'align': 'left'}

    // text position moves relative of the previous text
    doc.font(opts.font.weight.boldItalic)
    doc.fontSize(opts.font.size.normal)
    lineHeight = doc.currentLineHeight()

    curr_x = opts.page.margin.left + 10
    curr_y = opts.page.margin.top + (lineHeight*1.25)

    doc.text('Which of the following do you think a financial planner can help you with? (% Yes)', curr_x, curr_y, {'width': box_par_width, 'align': 'justify'});

    // track the y-coords for table heading
    // the table headers
    thead_heading = [];
    thead_heading[0] = {'width': 0.55, 'data': ''};
    thead_heading[1] = {'width': 0.15, 'data': 'Your Shopper'};
    thead_heading[2] = {'width': 0.15, 'data': licensee_name};
    thead_heading[3] = {'width': 0.15, 'data': 'Industry Average'};

    // resets back to pointy edge
    doc.lineWidth(0.3)
    // lineJoin() cause a subtle bug
    doc.lineCap('square');

    // track current width for the selected table column
    curr_x = opts.page.margin.left + 10
    curr_y = doc.y + (lineHeight*2)

    // draw the table headings
    doc.font(opts.font.weight.bold)
    doc.fontSize(opts.font.size.normal)

    lineHeight = doc.currentLineHeight()
    writable_width = (totalWidth*thead_heading[2].width) - 5
    doc_width = doc.widthOfString(thead_heading[2].data)
    answer_lines = lines_func(doc_width, writable_width)

    writable_width = (totalWidth*thead_heading[3].width) - 5
    doc_width = doc.widthOfString('Industry Average')
    licensee_lines = lines_func(doc_width, writable_width)

    max_line = Math.max(answer_lines, licensee_lines) + 1
    box_height = lineHeight*max_line

    for(var i=0; i<thead_heading.length; i++){
      if(i == 2){
        padding = ((max_line - answer_lines)*lineHeight) * 0.50
      }else{
        padding = lineHeight * 0.40
      }

      doc.rect(curr_x, curr_y, (totalWidth*thead_heading[i].width), box_height)
      doc.fillAndStroke(header_fill_style, '#000')

      doc.fillColor(table_header_color)
      doc.text(thead_heading[i].data, (curr_x + 2.5), (curr_y + padding), {'align': 'center', 'width': (totalWidth*thead_heading[i].width) - 5})

      curr_x = curr_x + (totalWidth*thead_heading[i].width)
    }

    doc.font(opts.font.weight.normal)

    // resets back to orig x-coord
    curr_x = opts.page.margin.left + 10

    // draw table rows
    // iterate the table row headings
    // add padding between rect box boundary and the text
    // each row is positioned using tableOpts['trow_height']*(i+1)
    curr_y = curr_y + box_height
    lineHeight = doc.currentLineHeight()

    for(var i = 0; i < views_planner_offer_questions.length; i++){
      var style = (i % 2) == 1 ? '#fafafa' : '#dedede';
      // set box_height according to font weight
      doc.font(opts.font.weight.boldItalic)

      switch(i){
        case 0:
        case 5:
          box_height = lineHeight*2;
          break;
        case 6:
          box_height = lineHeight*4;
          break;
        default:
          box_height = lineHeight*3;
          break;
      }

      // draw the bounding rect first
      for(var j=0; j<thead_heading.length; j++){
        doc.rect(curr_x, curr_y, (totalWidth*thead_heading[j].width), box_height)
        doc.fillAndStroke(style, '#000')

        curr_x += (totalWidth*thead_heading[j].width)
      }

      // resets back to orig x-coord
      curr_x = opts.page.margin.left + 10

      // adjust the padding between text and rect box
      var padding = 0;
      if(box_height > (lineHeight*2)){
        padding = box_height*0.30
      }else{
        padding = lineHeight*0.45
      }

      // @TODO refactor this, so many special cases
      if(i == 6){
        padding = lineHeight*1.50
      }

      // draw the text within the bounding rect
      doc.fillColor('#000')
      doc.text(views_planner_offer_questions[i].question, (curr_x + 2.5), (curr_y+(lineHeight*0.45)), {'width': (totalWidth*thead_heading[0].width) - 5, 'align': 'left'});
      doc.font(opts.font.weight.normal)

      // draw the self rating
      curr_x = curr_x + (totalWidth*thead_heading[0].width)
      doc.text(views_planner_offer_questions[i].answer, (curr_x + 2.5), (curr_y + padding), {'width': (totalWidth*thead_heading[1].width) - 5, 'align': 'center'});

      // draw licensee average
      curr_x = curr_x + (totalWidth*thead_heading[1].width)
      doc.text(views_planner_offer_questions[i].licensee, (curr_x + 2.5), (curr_y + padding), {'width': (totalWidth*thead_heading[2].width) - 5, 'align': 'center'});

      // draw industry average
      curr_x = curr_x + (totalWidth*thead_heading[2].width)
      doc.text(views_planner_offer_questions[i].industry, (curr_x + 2.5), (curr_y + padding), {'width': (totalWidth*thead_heading[3].width) - 5, 'align': 'center'});

      // reset x-coord to the leftmost writable page of the page
      curr_x = opts.page.margin.left + 10
      curr_y = curr_y + box_height
    }

    // draw BenefFA question code
    curr_x = opts.page.margin.left + 10
    curr_y = doc.y + box_height + lineHeight;

    doc.font(opts.font.weight.boldItalic)
    doc.fontSize(opts.font.size.normal)
    doc.text(BenefFA.question, curr_x, curr_y, parOpts)

    doc.font(opts.font.weight.italic)
    lineHeight = doc.currentLineHeight()

    curr_y = doc.y + lineHeight
    answer = '\“' + BenefFA.answer + '\”'
    doc.text(answer, curr_x, curr_y, parOpts)

   //doc.flushPages()


    // Assurance page
    doc.addPage(pageOptions)

    curr_x = opts.page.margin.left;
    curr_y = opts.page.margin.top;

    box_par_width = opts.page.width - (opts.page.margin.left + opts.page.margin.right)
    parOpts = {'width': box_par_width, 'align': 'left'}

    doc.fontSize(20)
    doc.font(opts.font.weight.bold)
    doc.text('ASSURANCES', curr_x, curr_y, parOpts)

    doc.fontSize(opts.font.size.normal)
    doc.font(opts.font.weight.normal)
    lineHeight = doc.currentLineHeight()

    curr_y = doc.y + lineHeight

    data = 'Assurances measures the adviser’s ability to demonstrate and communicate knowledge and skills.' 

    doc.font(opts.font.weight.bold)
    doc.text(data.slice(0, 10), curr_x, curr_y, {'width': box_par_width, 'align': 'left', 'continued': true})
    doc.font(opts.font.weight.normal)
    doc.text(data.slice(10))

    curr_y = doc.y + lineHeight

    doc.text('This is important as potential clients want to be sure they are trusting their financial future to someone with the necessary skills and experience to help them.', curr_x, curr_y, parOpts)

    doc.lineWidth(0.3)

    curr_x = opts.page.margin.left
    curr_y = doc.y + lineHeight
    box_height = lineHeight*2

    table_heading = ['Your score', licensee_name, 'Industry Average']
    // this should load the data
    table_row = []
    table_row[0] = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 37, 'r': Row}) ].v )
    table_row[1] = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 121, 'r': Row}) ].v )
    table_row[2] = '74'

    doc.font(opts.font.weight.bold)
    for(var i = 0; i<table_heading.length; i++){
      doc.rect(curr_x, curr_y, (box_par_width/3), box_height)
      doc.fillAndStroke(header_fill_style, '#000')
      doc.fillColor(table_header_color)
      doc.text(table_heading[i], curr_x, (curr_y+(box_height*0.30)), {'width': (box_par_width/3), 'align': 'center'})
      curr_x = curr_x + (box_par_width/3)
    }

    // position relative to the table header
    curr_y = curr_y + box_height

    doc.fontSize(36)
    lineHeight = doc.currentLineHeight()
    box_height = lineHeight*2

    curr_x = opts.page.margin.left
    for(var i = 0; i<table_row.length; i++){
      doc.rect(curr_x, curr_y, (box_par_width/3), box_height)
      doc.fillAndStroke('#fff', '#000')
      doc.fillColor('#000')
      doc.text(table_row[i], curr_x, (curr_y+(lineHeight*0.50)), {'width': (box_par_width/3), 'align': 'center'})

      curr_x = curr_x + (box_par_width/3)
    }

    doc.fontSize(opts.font.size.normal)
    doc.font(opts.font.weight.bold)
    lineHeight = doc.currentLineHeight()

    curr_x = opts.page.margin.left
    curr_y = doc.y + (lineHeight*4)
    doc.text('The shopper was asked to rate the following:', curr_x, curr_y, parOpts)

    curr_y = doc.y + lineHeight

    doc.font(opts.font.weight.normal)

    // 28 = AbilityDemo, 29 = AbilityExp
    for(var i=28; i<=32; i++){
      // exclude Impress column from iteration
      if(i != 30){
        q_key = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 0}) ].v
//        question = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 1}) ].v.replace(' (Answer)', '').replace(' (Text)', '').replace(' (Value)', '')
        switch(i){
          case 28:
            question = 'The planner\’s ability to demonstrate comprehensive knowledge of the company’s advice strategies and investment options for your situation.'
            break;
          case 29:
            question = 'The planner\’s ability to demonstrate expertise in financial planning.'
            break;
          case 31:
            question = 'The planner was clear and easy to understand.'
            break;
          case 32:
            question = 'The planner\’s ability to demonstrate effectively what benefits their service or advice will bring to you.'
            break;
        }

        doc.font(opts.font.weight.boldItalic)
        doc.text(question, curr_x, curr_y, {'align': 'left', 'width': box_par_width})

        curr_y = doc.y
        doc.image('resources/' + output_dir + '/' + q_key + '.png', curr_x, curr_y, {'scale': 0.70})

        // images are 50px tall
        curr_y = curr_y + (lineHeight*0.45) + 50
      }
    }

   //doc.flushPages()


    // Assurance page 2
    doc.addPage(pageOptions)

    curr_x = opts.page.margin.left;
    curr_y = opts.page.margin.top;

    box_par_width = opts.page.width - (opts.page.margin.left + opts.page.margin.right)
    parOpts = {'width': box_par_width, 'align': 'left'}

    doc.fontSize(opts.font.size.normal)
    doc.font(opts.font.weight.normal)
    lineHeight = doc.currentLineHeight()

    // @TODO, should refactor this segment to be consistent with others
    // CredPlanner
    question = worksheet[ XLSX.utils.encode_cell({'c': 26, 'r': 1}) ].v.replace(' (Text)', '')
    answer = worksheet[ XLSX.utils.encode_cell({'c': 26, 'r': Row}) ].v
    answer = '\“' + answer + '\”'

    // for review
    doc.font(opts.font.weight.boldItalic)
    doc.fontSize(opts.font.size.normal)
    doc.text(question, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

    curr_y = doc.y + lineHeight
    doc.font(opts.font.weight.italic)
    doc.fontSize(11)
    doc.text(answer, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

    curr_y = doc.y + (lineHeight*1.25)

    // Quali
    question = worksheet[ XLSX.utils.encode_cell({'c': 27, 'r': 1}) ].v.replace(' (Text)', '')
    answer = worksheet[ XLSX.utils.encode_cell({'c': 27, 'r': Row}) ].v
    answer = '\“' + answer + '\”'

    doc.font(opts.font.weight.boldItalic)
    doc.fontSize(opts.font.size.normal)
    doc.text(question, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

    curr_y = doc.y + lineHeight
    doc.font(opts.font.weight.italic)
    doc.fontSize(11)
    doc.text(answer, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

    curr_y = doc.y + (lineHeight*1.25)

    // Impress
    doc.fontSize(opts.font.size.normal)
    question = worksheet[ XLSX.utils.encode_cell({'c': 30, 'r': 1}) ].v.replace(' (Text)', '')
    answer = worksheet[ XLSX.utils.encode_cell({'c': 30, 'r': Row}) ].v
    answer = '\“' + answer + '\”'

    doc.font(opts.font.weight.boldItalic)
    doc.text(question, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

    curr_y = doc.y + lineHeight
    doc.font(opts.font.weight.italic)
    doc.fontSize(11)
    doc.text(answer, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})



//---------------------
    doc.fontSize(opts.font.size.normal)
    doc.font(opts.font.weight.boldItalic)
    lineHeight = doc.currentLineHeight()

    curr_x = opts.page.margin.left
    curr_y = doc.y + (lineHeight*2)
    doc.text('How much did the planner discuss the following during the meeting?', curr_x, curr_y, parOpts)

    curr_y = doc.y + lineHeight
//--
    // track the y-coords for table heading
    // the table headers
    thead_heading = [];
    thead_heading[0] = {'width': 0.46, 'data': ''};
    thead_heading[1] = {'width': 0.18, 'data': 'Your Shopper'};
    thead_heading[2] = {'width': 0.18, 'data': licensee_name};
    thead_heading[3] = {'width': 0.18, 'data': 'Industry Average'};

    // resets back to pointy edge
    doc.lineWidth(0.3)
    // lineJoin() cause a subtle bug
    doc.lineCap('square');
//--
    // draw the table headings
    doc.font(opts.font.weight.bold)
    doc.fontSize(opts.font.size.normal)

    lineHeight = doc.currentLineHeight()

    writable_width = (totalWidth*thead_heading[2].width) - 5
    doc_width = doc.widthOfString(thead_heading[2].data)
    answer_lines = lines_func(doc_width, writable_width)

    writable_width = (totalWidth*thead_heading[3].width) - 5
    doc_width = doc.widthOfString('Industry Average')
    licensee_lines = lines_func(doc_width, writable_width)

    max_line = Math.max(answer_lines, licensee_lines) + 1
    box_height = lineHeight*max_line

    curr_y = doc.y + lineHeight;
    for(var i=0; i<thead_heading.length; i++){
      if(i == 2){
        padding = ((max_line - answer_lines) * lineHeight) * 0.50
      }else{
        padding = lineHeight * 0.40
      }

      doc.rect(curr_x, curr_y, (totalWidth*thead_heading[i].width), box_height)
      doc.fillAndStroke(header_fill_style, '#000')

      doc.fillColor(table_header_color)
      doc.text(thead_heading[i].data, (curr_x + 2.5), (curr_y + padding), {'align': 'center', 'width': (totalWidth*thead_heading[i].width) - 5})

      curr_x = curr_x + (totalWidth*thead_heading[i].width)
    }

    doc.font(opts.font.weight.normal)

    // resets back to orig x-coord
    curr_x = opts.page.margin.left;

    // draw table rows
    // iterate the table row headings
    // add padding between rect box boundary and the text
    // each row is positioned using tableOpts['trow_height']*(i+1)

    curr_y = curr_y + box_height

    // reset box_height for table rows
    box_height = (lineHeight*3)

    // 33 = ExperV, 34 = QualiV, 35 = ProdV, 36 = SrvcsV
    scoreVal = []

//    question = worksheet[ XLSX.utils.encode_cell({'c': 33, 'r': 1}) ].v.replace(' (Text)', '')
    answer = worksheet[ XLSX.utils.encode_cell({'c': 33, 'r': Row}) ].v
    licensee_ave = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 117, 'r': Row}) ].v ) + '%'
    scoreVal[0] = {'key': 'ExperV', 'question': 'Their experience', 'answer': answer, 'licensee': licensee_ave, 'industry': '77%'}

//    question = worksheet[ XLSX.utils.encode_cell({'c': 34, 'r': 1}) ].v
//    question = question.replace(' (Value)', '').replace(' (Answer)', '').replace(' (Value)', '')
    answer = worksheet[ XLSX.utils.encode_cell({'c': 34, 'r': Row}) ].v
    licensee_ave = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 118, 'r': Row}) ].v ) + '%'
    scoreVal[1] = {'key': 'QualiV', 'question': 'Their qualifications', 'answer': answer, 'licensee': licensee_ave, 'industry': '63%'}

//    question = worksheet[ XLSX.utils.encode_cell({'c': 35, 'r': 1}) ].v
//    question = question.replace(' (Value)', '').replace(' (Answer)', '').replace(' (Value)', '')
    answer = worksheet[ XLSX.utils.encode_cell({'c': 35, 'r': Row}) ].v
    licensee_ave = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 119, 'r': Row}) ].v ) + '%'
    scoreVal[2] = {'key': 'ProdV', 'question': 'Their products', 'answer': answer, 'licensee': licensee_ave, 'industry': '81%'}

//    question = worksheet[ XLSX.utils.encode_cell({'c': 36, 'r': 1}) ].v
//    question = question.replace(' (Value)', '').replace(' (Answer)', '').replace(' (Value)', '')
    answer = worksheet[ XLSX.utils.encode_cell({'c': 36, 'r': Row}) ].v
    licensee_ave = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 120, 'r': Row}) ].v ) + '%'
    scoreVal[3] = {'key': 'SrvcsV', 'question': 'Their services', 'answer': answer, 'licensee': licensee_ave, 'industry': '87%'}

    for(var i = 0; i < scoreVal.length; i++){
      var style = (i % 2) == 1 ? '#fafafa' : '#dedede';

      // draw the bounding rect first
      for(var j=0; j<thead_heading.length; j++){
        doc.rect(curr_x, curr_y, (totalWidth*thead_heading[j].width), box_height)
        doc.fillAndStroke(style, '#000')

        curr_x += (totalWidth*thead_heading[j].width)
      }

      // resets back to orig x-coord
      curr_x = opts.page.margin.left

      // draw the text within the bounding rect
      doc.fillColor('#000')

      padding = lineHeight*0.45
      doc.font(opts.font.weight.boldItalic)
      doc.text(scoreVal[i].question, (curr_x + 5), (curr_y + lineHeight), {'width': (totalWidth*thead_heading[0].width), 'align': 'left'});
      doc.font(opts.font.weight.normal)

      // draw the self rating
      curr_x = curr_x + (totalWidth*thead_heading[0].width)
      doc.text(scoreVal[i].answer, curr_x, (curr_y + lineHeight), {'width': (totalWidth*thead_heading[1].width), 'align': 'center'})
      doc.font(opts.font.weight.normal)

      // draw licensee average
      curr_x = curr_x+(totalWidth*thead_heading[1].width)
      doc.text(scoreVal[i].licensee + ' said just right', (curr_x + 5), (curr_y + padding), {'width': (totalWidth*thead_heading[2].width) - 10, 'align': 'center'});

      // draw industry average
      curr_x = curr_x+(totalWidth*thead_heading[2].width)
      doc.text(scoreVal[i].industry + ' said just right', (curr_x + 5), (curr_y + padding), {'width': (totalWidth*thead_heading[3].width) - 10, 'align': 'center'});

      // reset x-coord to the leftmost writable page of the page
      curr_x = opts.page.margin.left;
      curr_y = curr_y + box_height
    }
//--

   //doc.flushPages()

    // Compliance page
    doc.addPage(pageOptions)

    curr_x = opts.page.margin.left;
    curr_y = opts.page.margin.top;

    box_par_width = opts.page.width - (opts.page.margin.left + opts.page.margin.right)
    parOpts = {'width': box_par_width, 'align': 'left'}

    doc.fontSize(20)
    doc.font(opts.font.weight.bold)
    doc.text('COMPLIANCE', curr_x, curr_y, parOpts)

    doc.fontSize(opts.font.size.normal)
    doc.font(opts.font.weight.normal)
    lineHeight = doc.currentLineHeight()

    curr_y = doc.y + lineHeight

    data = 'Compliance measures an adviser\'s ability to satisfy relevant financial regulations.' 

    doc.font(opts.font.weight.bold)
    doc.text(data.slice(0, 10), curr_x, curr_y, {'width': box_par_width, 'align': 'left', 'continued': true})
    doc.font(opts.font.weight.normal)
    doc.text(data.slice(10))

    doc.lineWidth(0.3)

    curr_x = opts.page.margin.left
    curr_y = doc.y + lineHeight
    box_height = lineHeight*2

    table_heading = ['Your score', licensee_name, 'Industry Average']
    // this should load the data
    table_row = []
    table_row[0] = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 46, 'r': Row}) ].v )
    table_row[1] = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 122, 'r': Row}) ].v )
    table_row[2] = '75'

    doc.font(opts.font.weight.bold)
    for(var i = 0; i<table_heading.length; i++){
      doc.rect(curr_x, curr_y, (totalWidth/3), box_height)
      doc.fillAndStroke(header_fill_style, '#000')
      doc.fillColor(table_header_color)
      doc.text(table_heading[i], curr_x, (curr_y+(box_height*0.30)), {'width': (totalWidth/3), 'align': 'center'})
      curr_x = curr_x + (totalWidth/3)
    }

    // position relative to the table header
    curr_y = curr_y + box_height

    doc.fontSize(36)
    lineHeight = doc.currentLineHeight()
    box_height = lineHeight*2

    curr_x = opts.page.margin.left
    for(var i = 0; i<table_row.length; i++){
      doc.rect(curr_x, curr_y, (totalWidth/3), box_height)
      doc.fillAndStroke('#fff', '#000')
      doc.fillColor('#000')
      doc.text(table_row[i], curr_x, (curr_y+(lineHeight*0.50)), {'width': (totalWidth/3), 'align': 'center'})

      curr_x = curr_x + (totalWidth/3)
    }

    doc.fontSize(opts.font.size.normal)
    doc.font(opts.font.weight.bold)
    lineHeight = doc.currentLineHeight()

    curr_x = opts.page.margin.left
    curr_y = doc.y + (lineHeight*4)
    doc.text('The shopper was asked:', curr_x, curr_y, parOpts)

    curr_y = doc.y + lineHeight

    // track the y-coords for table heading
    // the table headers
    var thead_heading = [];
    thead_heading[0] = {'width': 0.49, 'data': ''};
    thead_heading[1] = {'width': 0.17, 'data': 'Your Shopper'};
    thead_heading[2] = {'width': 0.17, 'data': licensee_name};
    thead_heading[3] = {'width': 0.17, 'data': 'Industry Average'};

    // resets back to pointy edge
    doc.lineWidth(0.3)
    // lineJoin() cause a subtle bug
    doc.lineCap('square');
//--
    // draw the table headings
    doc.font(opts.font.weight.bold)
    doc.fontSize(opts.font.size.normal)

    lineHeight = doc.currentLineHeight()

    writable_width = (totalWidth*thead_heading[2].width) - 5
    doc_width = doc.widthOfString(thead_heading[2].data)
    answer_lines = lines_func(doc_width, writable_width)

    writable_width = (totalWidth*thead_heading[3].width) - 5
    doc_width = doc.widthOfString('Industry Average')
    licensee_lines = lines_func(doc_width, writable_width)

    max_line = Math.max(answer_lines, licensee_lines) + 1
    box_height = lineHeight*max_line

    curr_y = doc.y + lineHeight;
    for(var i=0; i<thead_heading.length; i++){
      if(i == 2){
        padding = ((max_line - answer_lines) * lineHeight) * 0.50
      }else{
        padding = lineHeight * 0.40
      }

      doc.rect(curr_x, curr_y, (totalWidth*thead_heading[i].width), box_height)
      doc.fillAndStroke(header_fill_style, '#000')

      doc.fillColor(table_header_color)
      doc.text(thead_heading[i].data, (curr_x + 2.5), (curr_y + padding), {'align': 'center', 'width': (totalWidth*thead_heading[i].width) - 5})

      curr_x = curr_x + (totalWidth*thead_heading[i].width)
    }

    doc.font(opts.font.weight.normal)

    // resets back to orig x-coord
    curr_x = opts.page.margin.left;

    // draw table rows
    // iterate the table row headings
    // add padding between rect box boundary and the text
    // each row is positioned using tableOpts['trow_height']*(i+1)

    curr_y = curr_y + box_height

    // reset box_height for table rows
    box_height = (lineHeight*3)

    // 38 = RiskAtt, 41 = ShowFSG, 42 = ExplFSG, 43 = PrivyIssue, 44 = ExplIssue, 45 = SrvcProdOr
    // @todo put it in the top
    scoreVal = []

//    question = worksheet[ XLSX.utils.encode_cell({'c': 38, 'r': 1}) ].v.replace(' (Text)', '')
    answer = worksheet[ XLSX.utils.encode_cell({'c': 38, 'r': Row}) ].v
    licensee_ave = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 123, 'r': Row}) ].v ) + '%'
    scoreVal[0] = {'key': 'RiskAtt', 'question': 'Did the planner ask about your attitude to risk?', 'answer': answer, 'licensee': licensee_ave, 'industry': '77%'}

//    question = worksheet[ XLSX.utils.encode_cell({'c': 41, 'r': 1}) ].v
//    question = question.replace(' (Value)', '').replace(' (Answer)', '').replace(' (Value)', '')
    answer = worksheet[ XLSX.utils.encode_cell({'c': 41, 'r': Row}) ].v
    licensee_ave = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 124, 'r': Row}) ].v ) + '%'
    scoreVal[1] = {'key': 'ShowFSG', 'question': 'Did the planner give you or show you a financial services guide (FSG)?', 'answer': answer, 'licensee': licensee_ave, 'industry': '84%'}

//    question = worksheet[ XLSX.utils.encode_cell({'c': 42, 'r': 1}) ].v
//    question = question.replace(' (Value)', '').replace(' (Answer)', '').replace(' (Value)', '')
    answer = worksheet[ XLSX.utils.encode_cell({'c': 42, 'r': Row}) ].v
    licensee_ave = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 125, 'r': Row}) ].v ) + '%'
    scoreVal[2] = {'key': 'ExplFSG', 'question': 'Did the planner explain why the FSG was necessary?', 'answer': answer, 'licensee': licensee_ave, 'industry': '80%'}

//    question = worksheet[ XLSX.utils.encode_cell({'c': 43, 'r': 1}) ].v
//    question = question.replace(' (Value)', '').replace(' (Answer)', '').replace(' (Value)', '')
    answer = worksheet[ XLSX.utils.encode_cell({'c': 43, 'r': Row}) ].v
    licensee_ave = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 126, 'r': Row}) ].v ) + '%'
    scoreVal[3] = {'key': 'PrivyIssue', 'question': 'Did the planner discuss privacy issues with you?', 'answer': answer, 'licensee': licensee_ave, 'industry': '59%'}

//    question = worksheet[ XLSX.utils.encode_cell({'c': 44, 'r': 1}) ].v.replace(' (Answer)', '')
    answer = worksheet[ XLSX.utils.encode_cell({'c': 44, 'r': Row}) ].v
    licensee_ave = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 127, 'r': Row}) ].v ) + '%'
    scoreVal[4] = {'key': 'ExplIssue', 'question': 'Did the planner explain why it was necessary to discuss privacy issues?', 'answer': answer, 'licensee': licensee_ave, 'industry': '78%'}

//    question = worksheet[ XLSX.utils.encode_cell({'c': 45, 'r': 1}) ].v.replace(' (Text)', '')
    answer = worksheet[ XLSX.utils.encode_cell({'c': 45, 'r': Row}) ].v + ''
    licensee_ave = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 128, 'r': Row}) ].v ) + ''
    scoreVal[5] = {'key': 'SrvcProdOr', 'question': 'Did you feel the planner was more service or product focused?', 'answer': answer, 'licensee': licensee_ave, 'industry': '76'}

    for(var i = 0; i < scoreVal.length; i++){
      var max_line
      var style = (i % 2) == 1 ? '#fafafa' : '#dedede';

      writable_width = (totalWidth*thead_heading[0].width) - 10
      doc.fontSize(opts.font.size.normal)
      doc.font(opts.font.weight.bold)

      doc_width = doc.widthOfString(scoreVal[i].question)
      question_lines = Math.ceil(doc_width/writable_width)

      doc.fontSize(opts.font.size.normal)
      doc.font(opts.font.weight.normal)
      writable_width = (totalWidth*thead_heading[1].width) - 10
      doc_width = doc.widthOfString(scoreVal[i].answer)
      answer_lines = Math.ceil(doc_width/writable_width)

      writable_width = (totalWidth*thead_heading[2].width) - 10
      doc_width = doc.widthOfString(scoreVal[i].licensee)
      licensee_lines = Math.ceil(doc_width/writable_width) + 1

      writable_width = (totalWidth*thead_heading[3].width) - 10
      doc_width = doc.widthOfString(scoreVal[i].industry)
      industry_lines = Math.ceil(doc_width/writable_width) + 1

      max_line = Math.max(question_lines, answer_lines, licensee_lines, industry_lines) + 1

      box_height = lineHeight*max_line

      // draw the bounding rect first
      for(var j=0; j<thead_heading.length; j++){
        doc.rect(curr_x, curr_y, (totalWidth*thead_heading[j].width), box_height)
        doc.fillAndStroke(style, '#000')

        curr_x += (totalWidth*thead_heading[j].width)
      }

      // resets back to orig x-coord
      curr_x = opts.page.margin.left

      // draw the text within the bounding rect
      doc.fillColor('#000')

      doc.font(opts.font.weight.boldItalic)
      padding = ((max_line - question_lines) * lineHeight) * 0.50
      doc.text(scoreVal[i].question, (curr_x + 2.5), (curr_y + padding), {'width': (totalWidth*thead_heading[0].width) - 5, 'align': 'left'});
      doc.font(opts.font.weight.normal)

      // draw the self rating
      curr_x = curr_x + (totalWidth*thead_heading[0].width)

      if((max_line - answer_lines) > 1){
        padding = ((max_line - answer_lines) * lineHeight) * 0.50
      }else{
        padding = ((max_line - answer_lines) * lineHeight) * 0.25
      }
      doc.text(scoreVal[i].answer, (curr_x + 2.5), (curr_y + padding), {'width': (totalWidth*thead_heading[1].width) - 5, 'align': 'center'});
      doc.font(opts.font.weight.normal)

      // draw licensee average
      curr_x = curr_x+(totalWidth*thead_heading[1].width)
      if(i == (scoreVal.length - 1)){
        data = scoreVal[i].licensee
        padding = ((max_line - 1) * lineHeight) * 0.50
      }else{
        data = scoreVal[i].licensee + ' said yes'
        padding = ((max_line - licensee_lines) * lineHeight) * 0.50
      }
      doc.text(data, (curr_x + 2.5), (curr_y + padding), {'width': (totalWidth*thead_heading[2].width) - 5, 'align': 'center'});

      // draw industry average
      curr_x = curr_x+(totalWidth*thead_heading[2].width)
      if(i == (scoreVal.length - 1)){
        data = scoreVal[i].industry
        padding = ((max_line - 1) * lineHeight) * 0.50
      }else{
        data = scoreVal[i].industry + ' said yes'
        padding = ((max_line - industry_lines) * lineHeight) * 0.50
      }
      doc.text(data, (curr_x + 2.5), (curr_y + padding), {'width': (totalWidth*thead_heading[3].width) - 5, 'align': 'center'});

      // reset x-coord to the leftmost writable page of the page
      curr_x = opts.page.margin.left;
      curr_y = curr_y + box_height
    }

   //doc.flushPages()


    // Compliance page 2
    doc.addPage(pageOptions)

    curr_x = opts.page.margin.left;
    curr_y = opts.page.margin.top;

    box_par_width = opts.page.width - (opts.page.margin.left + opts.page.margin.right)
    parOpts = {'width': box_par_width, 'align': 'left'}

    doc.fontSize(opts.font.size.normal)
    doc.font(opts.font.weight.normal)
    lineHeight = doc.currentLineHeight()

    // 39 = DisclPay, 40 = DisclFees
    for(var i=39; i<=40; i++){
      q_key = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 0}) ].v
//      question = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 1}) ].v.replace(' (Answer)', '').replace(' (Text)', '')
      answer = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': Row}) ].v
      switch(i){
        case 39:
          question = 'Did the planner disclose how they were paid? (i.e salary, commission, fee for service, etc.)'
        break;
        case 40:
          question = 'Did the planner disclose the extent of any fees to be charged to you?'
        break;
        default:
          question = ''
        break;
      }

      doc.font(opts.font.weight.boldItalic)
      doc.text(question, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})
      doc.font(opts.font.weight.normal)
      //get max width for the questions
      max_width = doc.widthOfString('Your Shopper: ') + 30
      curr_y = doc.y + lineHeight

      doc.text('Your Shopper:', curr_x, curr_y, {'width': max_width, 'align': 'left'})
      doc.font(opts.font.weight.bold)
      doc.text(answer, (doc.x + max_width), curr_y, {'width': (box_par_width - max_width), 'align': 'left'})

      curr_y = doc.y + lineHeight
      doc.font(opts.font.weight.normal)
      doc.image(licensee_images + '/' + q_key + '.png', curr_x, curr_y, {'scale': 0.1275})

      curr_y = doc.y + (lineHeight*13);
    }

   //doc.flushPages()

    // Quality page
    doc.addPage(pageOptions)

    curr_x = opts.page.margin.left;
    curr_y = opts.page.margin.top;

    box_par_width = opts.page.width - (opts.page.margin.left + opts.page.margin.right)
    parOpts = {'width': box_par_width, 'align': 'left'}

    doc.fontSize(20)
    doc.font(opts.font.weight.bold)
    doc.text('QUALITY', curr_x, curr_y, parOpts)

    doc.fontSize(opts.font.size.normal)
    doc.font(opts.font.weight.normal)
    lineHeight = doc.currentLineHeight()

    curr_y = doc.y + lineHeight

    data = 'Quality is a measure of an adviser’s ability to satisfy customer needs and provide perceived value.' 

    doc.font(opts.font.weight.bold)
    doc.text(data.slice(0, 7), curr_x, curr_y, {'width': box_par_width, 'align': 'left', 'continued': true})
    doc.font(opts.font.weight.normal)
    doc.text(data.slice(7))

    doc.lineWidth(0.3)

    curr_x = opts.page.margin.left
    curr_y = doc.y + lineHeight
    box_height = lineHeight*2

    table_heading = ['Your score', licensee_name, 'Industry Average']
    // this should load the data
    table_row = []
    table_row[0] = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 50, 'r': Row}) ].v )
    table_row[1] = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 129, 'r': Row}) ].v )
    table_row[2] = '68'

    doc.font(opts.font.weight.bold)
    for(var i = 0; i<table_heading.length; i++){
      doc.rect(curr_x, curr_y, (totalWidth/3), box_height)
      doc.fillAndStroke(header_fill_style, '#000')
      doc.fillColor(table_header_color)
      doc.text(table_heading[i], curr_x, (curr_y+(box_height*0.30)), {'width': (totalWidth/3), 'align': 'center'})
      curr_x = curr_x + (totalWidth/3)
    }

    // position relative to the table header
    curr_y = curr_y + box_height

    doc.fontSize(36)
    lineHeight = doc.currentLineHeight()
    box_height = lineHeight*2

    curr_x = opts.page.margin.left
    for(var i = 0; i<table_row.length; i++){
      doc.rect(curr_x, curr_y, (totalWidth/3), box_height)
      doc.fillAndStroke('#fff', '#000')
      doc.fillColor('#000')
      doc.text(table_row[i], (curr_x + 2.5), (curr_y+(lineHeight*0.50)), {'width': (totalWidth/3) - 5, 'align': 'center'})

      curr_x = curr_x + (totalWidth/3)
    }

    doc.fontSize(opts.font.size.normal)
    doc.font(opts.font.weight.normal)
    lineHeight = doc.currentLineHeight()

    curr_x = opts.page.margin.left
    curr_y = doc.y + (lineHeight*4)
    doc.font(opts.font.weight.bold)
    doc.text('The shopper was asked:', curr_x, curr_y, parOpts)

    curr_y = doc.y + lineHeight

    // 47 = ConvReco, 48 = FeesPay, 49 = PlanSrvcs
    doc.font(opts.font.weight.normal)
    for(var i=47; i<=49; i++){
      q_key = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 0}) ].v
      question = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 1}) ].v.replace(' (Answer)', '').replace(' (Text)', '').replace(' (Value)', '')

      doc.font(opts.font.weight.boldItalic)
      doc.text(question, curr_x, curr_y, {'align': 'left', 'width': box_par_width})

      curr_y = doc.y
      doc.image('resources/' + output_dir + '/' + q_key + '.png', curr_x, curr_y, {'scale': 0.70})

      // images are 50px tall
      curr_y = curr_y + (lineHeight*0.45) + 50
    }

   //doc.flushPages()


    // Understanding page
    doc.addPage(pageOptions)

    curr_x = opts.page.margin.left;
    curr_y = opts.page.margin.top;

    box_par_width = opts.page.width - (opts.page.margin.left + opts.page.margin.right)
    parOpts = {'width': box_par_width, 'align': 'left'}

    doc.fontSize(20)
    doc.font(opts.font.weight.bold)
    doc.text('UNDERSTANDING', curr_x, curr_y, parOpts)

    doc.fontSize(opts.font.size.normal)
    doc.font(opts.font.weight.normal)
    lineHeight = doc.currentLineHeight()

    curr_y = doc.y + lineHeight

    data = 'Understanding measures an adviser\'s ability to understand client needs.' 

    doc.font(opts.font.weight.bold)
    doc.text(data.slice(0, 13), curr_x, curr_y, {'width': box_par_width, 'align': 'left', 'continued': true})
    doc.font(opts.font.weight.normal)
    doc.text(data.slice(13))

    doc.lineWidth(0.3)

    curr_x = opts.page.margin.left
    curr_y = doc.y + lineHeight
    box_height = lineHeight*2

    table_heading = ['Your score', licensee_name, 'Industry Average']
    // this should load the data
    table_row = []
    table_row[0] = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 56, 'r': Row}) ].v )
    table_row[1] = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 133, 'r': Row}) ].v )
    table_row[2] = '80'

    doc.font(opts.font.weight.bold)
    for(var i = 0; i<table_heading.length; i++){
      doc.rect(curr_x, curr_y, (totalWidth/3), box_height)
      doc.fillAndStroke(header_fill_style, '#000')
      doc.fillColor(table_header_color)
      doc.text(table_heading[i], (curr_x + 2.5), (curr_y+(box_height*0.30)), {'width': (totalWidth/3) - 5, 'align': 'center'})
      curr_x = curr_x + (totalWidth/3)
    }

    // position relative to the table header
    curr_y = curr_y + box_height

    doc.fontSize(36)
    lineHeight = doc.currentLineHeight()
    box_height = lineHeight*2

    curr_x = opts.page.margin.left
    for(var i = 0; i<table_row.length; i++){
      doc.rect(curr_x, curr_y, (totalWidth/3), box_height)
      doc.fillAndStroke('#fff', '#000')
      doc.fillColor('#000')
      doc.text(table_row[i], curr_x, (curr_y+(lineHeight*0.50)), {'width': (totalWidth/3), 'align': 'center'})

      curr_x = curr_x + (totalWidth/3)
    }

    doc.fontSize(opts.font.size.normal)
    doc.font(opts.font.weight.bold)
    lineHeight = doc.currentLineHeight()

    curr_x = opts.page.margin.left
    curr_y = doc.y + (lineHeight*3)

    doc.text('The shopper was asked:', curr_x, curr_y, parOpts)

    curr_y = doc.y + lineHeight

    // 51 = ListenSkill, 52 = Goals, 53 = DemoGoals, 54 = ReadFact, 55 = WellPrep
    doc.font(opts.font.weight.normal)
    for(var i=51; i<=55; i++){
      q_key = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 0}) ].v
//      question = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 1}) ].v.replace(' (Answer)', '').replace(' (Text)', '').replace(' (Value)', '')
      switch(i){
        case 51:
          question = 'How would you rate the planner’s listening skills?'
        break;
        case 52:
          question = 'How well did the planner take the time to understand your needs and goals?'
        break;
        case 53:
          question = 'How well did the planner demonstrate that they understood your personal needs and goals?'
        break;
        case 54:
          question = 'Do you think the planner briefed themselves on any information you provided to them?'
        break;
        case 55:
          question = 'How well prepared was the planner for the meeting?'
        break;
        default:
          question = ''
        break;
      }

      doc.font(opts.font.weight.boldItalic)
      doc.text(question, curr_x, curr_y, {'align': 'left', 'width': box_par_width})
      doc.font(opts.font.weight.normal)


      curr_y = doc.y
      doc.image('resources/' + output_dir + '/' + q_key + '.png', curr_x, curr_y, {'scale': 0.70})

      // images are 50px tall
      curr_y = curr_y + (lineHeight*0.45) + 50
    }

   //doc.flushPages()


    // Intention page
    doc.addPage(pageOptions)

    curr_x = opts.page.margin.left;
    curr_y = opts.page.margin.top;

    box_par_width = opts.page.width - (opts.page.margin.left + opts.page.margin.right)
    parOpts = {'width': box_par_width, 'align': 'left'}

    doc.fontSize(20)
    doc.font(opts.font.weight.bold)
    doc.text('INTENTION', curr_x, curr_y, parOpts)

    doc.fontSize(opts.font.size.normal)
    doc.font(opts.font.weight.normal)
    lineHeight = doc.currentLineHeight()

    curr_y = doc.y + lineHeight

    data = 'Intention measures the likelihood of a shopper using, reusing or recommending a planner.' 

    doc.font(opts.font.weight.bold)
    doc.text(data.slice(0, 9), curr_x, curr_y, {'width': box_par_width, 'align': 'left', 'continued': true})
    doc.font(opts.font.weight.normal)
    doc.text(data.slice(9))

    curr_y = doc.y + lineHeight

    data = 'This measure reflects the likelihood of a shopper:'

    //bulleted items
    curr_x = opts.page.margin.left + 18
    curr_y = doc.y + lineHeight
    doc.circle(curr_x, (curr_y + (lineHeight/2.5)), 2).fill('#000').stroke()
    doc.text('proceeding to a second meeting', (curr_x + 18), curr_y)

    curr_x = opts.page.margin.left + 18
    curr_y = doc.y
    doc.circle(curr_x, (curr_y + (lineHeight/2.5)), 2).fill('#000').stroke()
    doc.text('recommending the planner to others', (curr_x + 18), curr_y)

    curr_x = opts.page.margin.left;
    curr_y = doc.y + lineHeight

    doc.text('This category will reflect performance across the whole process as well as the effectiveness with which the planner has created a ‘call to action’.', curr_x, curr_y, parOpts)

    doc.lineWidth(0.3)

    curr_x = opts.page.margin.left
    curr_y = doc.y + lineHeight
    box_height = lineHeight*2

    table_heading = ['Your score', licensee_name, 'Industry Average']
    // this should load the data
    table_row = []
    table_row[0] = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 65, 'r': Row}) ].v )
    table_row[1] = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 139, 'r': Row}) ].v )
    table_row[2] = '62'

    doc.font(opts.font.weight.bold)
    for(var i = 0; i<table_heading.length; i++){
      doc.rect(curr_x, curr_y, (totalWidth/3), box_height)
      doc.fillAndStroke(header_fill_style, '#000')
      doc.fillColor(table_header_color)
      doc.text(table_heading[i], (curr_x + 2.5), (curr_y+(box_height*0.30)), {'width': (totalWidth/3) - 5, 'align': 'center'})
      curr_x = curr_x + (totalWidth/3)
    }

    // position relative to the table header
    curr_y = curr_y + box_height

    doc.fontSize(36)
    lineHeight = doc.currentLineHeight()
    box_height = lineHeight*2

    curr_x = opts.page.margin.left
    for(var i = 0; i<table_row.length; i++){
      doc.rect(curr_x, curr_y, (totalWidth/3), box_height)
      doc.fillAndStroke('#fff', '#000')
      doc.fillColor('#000')
      doc.text(table_row[i], curr_x, (curr_y+(lineHeight*0.50)), {'width': (totalWidth/3), 'align': 'center'})

      curr_x = curr_x + (totalWidth/3)
    }

    doc.fontSize(opts.font.size.normal)
    doc.font(opts.font.weight.bold)
    lineHeight = doc.currentLineHeight()

    curr_x = opts.page.margin.left
    curr_y = doc.y + (lineHeight*3)

    doc.text('The shopper was asked:', curr_x, curr_y, parOpts)

    curr_y = doc.y + lineHeight

// --
    // 57 = M_2ndMeet, 59 = RecoP
    for(var i=57; i<=59; i=i+2){
      q_key = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 0}) ].v
//      question = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 1}) ].v.replace(' (Answer)', '').replace(' (Text)', '').replace(' (Value)', '')
      switch(i){
        case 57:
          question = 'How likely are you to proceed to a second meeting?'
        break;
        case 59:
          question = 'How likely are you to recommend this planner to friends, family or colleagues?'
        break;
        default:
          question = ''
        break;
      }

      doc.font(opts.font.weight.boldItalic)
      doc.text(question, curr_x, curr_y, {'align': 'left', 'width': box_par_width})

      curr_y = doc.y
      doc.image('resources/' + output_dir + '/' + q_key + '.png', curr_x, curr_y, {'scale': 0.70})

      // images are 50px tall
      curr_y = curr_y + (lineHeight*0.45) + 50

      answer = worksheet[ XLSX.utils.encode_cell({'c': (i+1), 'r': Row}) ].v

      curr_y = doc.y + (lineHeight*1.25)

      doc.text('Reason:', curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

      curr_y = doc.y

      doc.font(opts.font.weight.italic)
      answer = '\“' + answer + '\”'
      doc.text(answer, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

      curr_y = doc.y + (lineHeight*1.25)

    }

// --

   //doc.flushPages()


    // Intention page 2
    doc.addPage(pageOptions)

    curr_x = opts.page.margin.left;
    curr_y = opts.page.margin.top;

    box_par_width = opts.page.width - (opts.page.margin.left + opts.page.margin.right)
    parOpts = {'width': box_par_width, 'align': 'left'}

    doc.fontSize(opts.font.size.normal)
    doc.font(opts.font.weight.normal)
    lineHeight = doc.currentLineHeight()

    curr_y = curr_y + lineHeight

    doc.lineWidth(0.3)

    // track the y-coords for table heading
    // the table headers
    thead_heading = [];
    thead_heading[0] = {'width': 0.49, 'data': ''};
    thead_heading[1] = {'width': 0.17, 'data': 'Your Shopper'};
    thead_heading[2] = {'width': 0.17, 'data': licensee_name};
    thead_heading[3] = {'width': 0.17, 'data': 'Industry Average'};

    // lineJoin() cause a subtle bug
    doc.lineCap('square');

    // draw the table headings
    doc.font(opts.font.weight.bold)
    doc.fontSize(opts.font.size.normal)

    lineHeight = doc.currentLineHeight()

    writable_width = (totalWidth*thead_heading[2].width) - 5
    doc_width = doc.widthOfString(thead_heading[2].data)
    answer_lines = lines_func(doc_width, writable_width)

    writable_width = (totalWidth*thead_heading[3].width) - 5
    doc_width = doc.widthOfString('Industry Average')
    licensee_lines = lines_func(doc_width, writable_width)

    max_line = Math.max(answer_lines, licensee_lines) + 1
    box_height = lineHeight*max_line

    curr_y = curr_y + lineHeight;
    for(var i=0; i<thead_heading.length; i++){

      if(i == 2){
        padding = ((max_line - answer_lines) * lineHeight) * 0.50
      }else{
        padding = lineHeight * 0.40
      }

      doc.rect(curr_x, curr_y, (totalWidth*thead_heading[i].width), box_height)
      doc.fillAndStroke(header_fill_style, '#000')

      doc.fillColor(table_header_color)
      doc.text(thead_heading[i].data, (curr_x + 2.5), (curr_y + padding), {'align': 'center', 'width': (totalWidth*thead_heading[i].width) - 5})

      curr_x = curr_x + (totalWidth*thead_heading[i].width)
    }

    doc.font(opts.font.weight.normal)

    // resets back to orig x-coord
    curr_x = opts.page.margin.left;

    // draw table rows
    // iterate the table row headings
    // add padding between rect box boundary and the text
    // each row is positioned using tableOpts['trow_height']*(i+1)

    curr_y = curr_y + box_height


    // 61 = ProcUse, 63 = ImprSit
    var k = 0;
    for(var i=61; i<= 63; i=i+2){
      var style = (k % 2) == 1 ? '#fafafa' : '#dedede';

      // @TODO refactor this table generation
      switch(i){
        case 61:
          question = 'Did you or will you proceed to use this planner (i.e. pay for advice)?'
          // @TODO refactor
          industry_ave = '14% said yes'
          licensee_ave = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 142, 'r': Row}) ].v ) + '%'

          // reset the box_height for table rows
          box_height = (lineHeight*4)
          padding = lineHeight * 1.50
          industry_padding = lineHeight

          break;
        case 63:
          question = 'Do you think this planner can improve your situation?'
          // @TODO refactor
          industry_ave = '54% said yes'
          licensee_ave = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 143, 'r': Row}) ].v ) + '%'

          // reset the box_height for table rows
          box_height = (lineHeight*3)
          padding = lineHeight
          industry_padding = lineHeight * 0.45

          break;
        default:
          question = ''
          break;
      }


      // draw the bounding rect first
      for(var j=0; j<thead_heading.length; j++){
        doc.rect(curr_x, curr_y, (totalWidth*thead_heading[j].width), box_height)
        doc.fillAndStroke('#fafafa', '#000')

        curr_x += (totalWidth*thead_heading[j].width)
      }

      q_key = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 0}) ].v
//      question = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 1}) ].v.replace(' (Answer)', '')
      answer = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': Row}) ].v

      curr_x = opts.page.margin.left;

      doc.fillColor('#000')
      // draw the text within the bounding rect
      doc.font(opts.font.weight.boldItalic)
      doc.text(question, (curr_x + 5), (curr_y + (lineHeight*0.45)), {'width': (totalWidth*thead_heading[0].width) - 10 , 'align': 'left'});
      doc.font(opts.font.weight.normal)

      // draw the self rating
      //@TODO refactor
      if(answer.toLowerCase() == 'yes, planning to proceed'){
        padding = lineHeight * 0.45
      }
      curr_x = curr_x + (totalWidth*thead_heading[0].width)
      doc.text(answer, (curr_x + 5), (curr_y + padding), {'width': (totalWidth*thead_heading[1].width) - 10, 'align': 'center'});
      doc.font(opts.font.weight.normal)

      // licensee average
      curr_x = curr_x + (totalWidth*thead_heading[1].width)
      doc.text(licensee_ave + ' said yes', (curr_x + 5), (curr_y + industry_padding), {'width': (totalWidth*thead_heading[2].width) - 10, 'align': 'center'});

      // draw industry average
      curr_x = curr_x + (totalWidth*thead_heading[2].width)

      // we can refator this call
      doc.text(industry_ave, (curr_x + 5), (curr_y + industry_padding), {'width': (totalWidth*thead_heading[3].width) - 10, 'align': 'center'});

      // reset x-coord to the leftmost writable page of the page
      curr_x = opts.page.margin.left;

      curr_y = curr_y + box_height

      // text explanation
      c_addr = XLSX.utils.encode_cell({'c': i+1, 'r': Row})
      if(worksheet[c_addr]){
        answer = worksheet[c_addr].v
      }else{
        answer = ''
      }

      // dynamically allocation of table cell height
      answer = '\“' + answer + '\”'
      doc_width = doc.widthOfString(answer)
      lines = lines_func(doc_width, (totalWidth - 25))
      box_height = doc.currentLineHeight() * (lines + 2)
      doc.rect(curr_x, curr_y, totalWidth, box_height)

      doc.fillAndStroke('#dedede', '#000')

      doc.fillColor('#000')

      switch(i){
        case 61:
          question = 'What were the critical factors that led to your decision?'
        break;
        case 63:
          question = 'Why do you say that?'
        break;
        default:
          question = ''
        break;
      }

      doc.font(opts.font.weight.boldItalic)
      doc.text(question, (curr_x + 5), (curr_y + (lineHeight*0.45)), {'width': totalWidth - 10, 'align': 'left'})

      doc.font(opts.font.weight.italic)
      doc.text(answer, (curr_x + 5), curr_y + (lineHeight*1.5), {'width': totalWidth - 10, 'align': 'left'})

      curr_x = opts.page.margin.left;
      curr_y = curr_y + box_height
      k++;
    }

   //doc.flushPages()


    // Reaction page
    doc.addPage(pageOptions)

    curr_x = opts.page.margin.left;
    curr_y = opts.page.margin.top;

    box_par_width = opts.page.width - (opts.page.margin.left + opts.page.margin.right)
    parOpts = {'width': box_par_width, 'align': 'left'}

    doc.fontSize(20)
    doc.font(opts.font.weight.bold)
    doc.text('REACTION', curr_x, curr_y, parOpts)

    doc.fontSize(opts.font.size.normal)
    doc.font(opts.font.weight.normal)
    lineHeight = doc.currentLineHeight()

    curr_y = doc.y + lineHeight

    data = 'Reaction measures a client’s emotive/affective response to the purchase process.' 

    doc.font(opts.font.weight.bold)
    doc.text(data.slice(0, 8), curr_x, curr_y, {'width': box_par_width, 'align': 'left', 'continued': true})
    doc.font(opts.font.weight.normal)
    doc.text(data.slice(8))

    curr_y = doc.y + lineHeight

    doc.text('This is an important category as historically many of the underlying dimensions correlate strongly with overall client satisfaction.', curr_x, curr_y, parOpts)

//--
    doc.lineWidth(0.3)

    curr_x = opts.page.margin.left
    curr_y = doc.y + lineHeight
    box_height = lineHeight*2

    table_heading = ['Your score', licensee_name, 'Industry Average']
    // this should load the data
    table_row = []
    table_row[0] = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 79, 'r': Row}) ].v )
    table_row[1] = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 144, 'r': Row}) ].v )
    table_row[2] = '84'

    doc.font(opts.font.weight.bold)
    for(var i = 0; i<table_heading.length; i++){
      doc.rect(curr_x, curr_y, (totalWidth/3), box_height)
      doc.fillAndStroke(header_fill_style, '#000')
      doc.fillColor(table_header_color)
      doc.text(table_heading[i], (curr_x + 2.5), (curr_y+(box_height*0.30)), {'width': (totalWidth/3) - 5, 'align': 'center'})
      curr_x = curr_x + (totalWidth/3)
    }

    // position relative to the table header
    curr_y = curr_y + box_height

    doc.fontSize(36)
    lineHeight = doc.currentLineHeight()
    box_height = lineHeight*2

    curr_x = opts.page.margin.left
    for(var i = 0; i<table_row.length; i++){
      doc.rect(curr_x, curr_y, (totalWidth/3), box_height)
      doc.fillAndStroke('#fff', '#000')
      doc.fillColor('#000')
      doc.text(table_row[i], curr_x, (curr_y+(lineHeight*0.50)), {'width': (totalWidth/3), 'align': 'center'})

      curr_x = curr_x + (totalWidth/3)
    }

    doc.fontSize(opts.font.size.normal)
    doc.font(opts.font.weight.bold)
    lineHeight = doc.currentLineHeight()

    curr_x = opts.page.margin.left
    curr_y = doc.y + (lineHeight*3)
    doc.text('The shopper was asked:', curr_x, curr_y, parOpts)

//    doc.font(opts.font.weight.normal)

    curr_y = doc.y + lineHeight

    q_key = worksheet[ XLSX.utils.encode_cell({'c': 66, 'r': 0}) ].v
//    question = worksheet[ XLSX.utils.encode_cell({'c': 66, 'r': 1}) ].v
//    question = question.replace(' (Value)', '').replace(' (Answer)', '')
    question = 'How well did the planner demonstrate their keenness for your business? (e.g. Was the planner just going through a process, or were they actively engaged in the process of helping meet your needs?)'

    answer = worksheet[ XLSX.utils.encode_cell({'c': 67, 'r': Row}) ].v

    doc.font(opts.font.weight.boldItalic)
    doc.text(question, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

    curr_y = doc.y
    doc.image('resources/' + output_dir + '/' + q_key + '.png', curr_x, curr_y, {'scale': 0.70})

    // images are 50px tall
    curr_y = curr_y + (lineHeight*0.45) + 50

    doc.text('Reason:', curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

    curr_y = doc.y

    doc.font(opts.font.weight.italic)
    answer = '\“' + answer + '\”'
    doc.text(answer, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

    curr_y = doc.y + lineHeight

    //68 = Gimpress, 69 = Influence, 70 = Enthuse
    for(var i=68; i<=70; i++){
      q_key = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 0}) ].v
//      question = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 1}) ].v.replace(' (Answer)', '').replace(' (Text)', '').replace(' (Value)', '')
      switch(i){
        case 68:
          question = 'Ability to create a good impression'
          break;
        case 69:
          question = 'Ability to influence you'
          break;
        case 70:
          question = 'Ability to enthuse you'
          break;
        default:
          question = ''
          break;
      }

      doc.font(opts.font.weight.boldItalic)
      doc.text(question, curr_x, curr_y, {'align': 'left', 'width': box_par_width})

      curr_y = doc.y
      doc.image('resources/' + output_dir + '/' + q_key + '.png', curr_x, curr_y, {'scale': 0.70})

      // images are 50px tall
      curr_y = curr_y + (lineHeight*0.45) + 40
    }

   //doc.flushPages()

    // Reaction page 2
    doc.addPage(pageOptions)

    curr_x = opts.page.margin.left;
    curr_y = opts.page.margin.top;

    box_par_width = opts.page.width - (opts.page.margin.left + opts.page.margin.right)
    parOpts = {'width': box_par_width, 'align': 'left'}

    doc.fontSize(opts.font.size.normal)
    doc.font(opts.font.weight.normal)
    lineHeight = doc.currentLineHeight()

    // 71 = Reltn, 72 = Rapprt
    for(var i=71; i<=72; i++){
      q_key = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 0}) ].v
//      question = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 1}) ].v.replace(' (Answer)', '').replace(' (Text)', '').replace(' (Value)', '')
      switch(i){
        case 71:
          question = 'Ability to build relationships'
          break;
        case 72:
          question = 'Rapport building skills'
          break;
        default:
          question = ''
          break;
      }

      doc.font(opts.font.weight.boldItalic)
      doc.text(question, curr_x, curr_y, {'align': 'left', 'width': box_par_width})
      doc.font(opts.font.weight.normal)

      curr_y = doc.y
      doc.image('resources/' + output_dir + '/' + q_key + '.png', curr_x, curr_y, {'scale': 0.70})

      // images are 50px tall
      curr_y = curr_y + (lineHeight*0.45) + 40
    }

    curr_y = doc.y + (lineHeight*2)

    doc.font(opts.font.weight.bold)
    doc.text('Additionally, the shopper was asked:', curr_x, curr_y, parOpts)

    curr_y = doc.y + lineHeight

    doc.lineWidth(0.3)

    // track the y-coords for table heading
    // the table headers
    var thead_heading = [];
    thead_heading[0] = {'width': 0.49, 'data': ''};
    thead_heading[1] = {'width': 0.17, 'data': 'Your Shopper'};
    thead_heading[2] = {'width': 0.17, 'data': licensee_name};
    thead_heading[3] = {'width': 0.17, 'data': 'Industry Average'};

    // lineJoin() cause a subtle bug
    doc.lineCap('square');

    // draw the table headings
    doc.font(opts.font.weight.bold)
    doc.fontSize(opts.font.size.normal)

    lineHeight = doc.currentLineHeight()

    writable_width = (totalWidth*thead_heading[2].width) - 5
    doc_width = doc.widthOfString(thead_heading[2].data)
    answer_lines = lines_func(doc_width, writable_width)

    writable_width = (totalWidth*thead_heading[3].width) - 5
    doc_width = doc.widthOfString('Industry Average')
    licensee_lines = lines_func(doc_width, writable_width)

    max_line = Math.max(answer_lines, licensee_lines) + 1
    box_height = lineHeight*max_line

    curr_y = doc.y + lineHeight;
    for(var i=0; i<thead_heading.length; i++){
      if(i == 2){
        padding = ((max_line - answer_lines) * lineHeight) * 0.50;
      }else{
        padding = lineHeight * 0.40;
      }

      doc.rect(curr_x, curr_y, (totalWidth*thead_heading[i].width), box_height)
      doc.fillAndStroke(header_fill_style, '#000')

      doc.fillColor(table_header_color)
      doc.text(thead_heading[i].data, (curr_x + 2.5), (curr_y + padding), {'align': 'center', 'width': (totalWidth*thead_heading[i].width) - 5})

      curr_x = curr_x + (totalWidth*thead_heading[i].width)
    }

    doc.font(opts.font.weight.normal)

    // resets back to orig x-coord
    curr_x = opts.page.margin.left;

    // draw table rows
    // iterate the table row headings
    // add padding between rect box boundary and the text
    // each row is positioned using tableOpts['trow_height']*(i+1)

    curr_y = curr_y + box_height

    // 73 = Probs, 75 = Honesty, 77 = Trust
    var k = 0;
    for(var i=73; i<= 77; i=i+2){
      var style = (k % 2) == 1 ? '#fafafa' : '#dedede';

      switch(i){
        case 73:
          question = 'Were there any problems during your conversation or contact with the planner?'
          // @TODO refactor
          industry_ave = '7% said yes'
          licensee_ave = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 151, 'r': Row}) ].v ) + '%'

          // reset the box_height for table rows
          box_height = (lineHeight*4)
          padding = lineHeight * 1.50

          break;
        case 75:
          question = 'Do you think they were honest throughout the meeting?'
          // @TODO refactor
          industry_ave = '92% said yes'
          licensee_ave = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 152, 'r': Row}) ].v ) + '%'

          // reset the box_height for table rows
          box_height = (lineHeight*3)
          padding = lineHeight

          break;
        case 77:
          question = 'How much do you trust the planner you met with?'
          // @TODO refactor
          industry_ave = '82'
          licensee_ave = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 153, 'r': Row}) ].v )

          // reset the box_height for table rows
          box_height = (lineHeight*3)
          padding = lineHeight

          break;
        default:
          question = ''
          break;
      }

      writable_width =  (totalWidth*thead_heading[0].width) - 10
      doc_width = doc.widthOfString(question)
      licensee_lines = lines_func(doc_width, writable_width)

      max_line = Math.max(answer_lines, licensee_lines) + 1
      box_height = lineHeight*max_line

      // draw the bounding rect first
      for(var j=0; j<thead_heading.length; j++){
        doc.rect(curr_x, curr_y, (totalWidth*thead_heading[j].width), box_height)
        doc.fillAndStroke('#fafafa', '#000')

        curr_x += (totalWidth*thead_heading[j].width)
      }

      q_key = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 0}) ].v
//      question = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 1}) ].v.replace(' (Answer)', '')
      answer = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': Row}) ].v


      curr_x = opts.page.margin.left

      doc.fillColor('#000')
      // draw the text within the bounding rect
      doc.font(opts.font.weight.boldItalic)
      doc.text(question, (curr_x + 5), (curr_y + (lineHeight*0.45)), {'width': (totalWidth*thead_heading[0].width) - 10, 'align': 'left'});
      doc.font(opts.font.weight.normal)

      // draw the self rating
      curr_x = curr_x + (totalWidth*thead_heading[0].width)
      doc.text(answer, (curr_x + 5), (curr_y + padding), {'width': (totalWidth*thead_heading[1].width) - 10, 'align': 'center'});
      doc.font(opts.font.weight.normal)

      // we can refator this call
      // this  should not be so special case
      if((i == 73) || (i == 77)){
        padding = lineHeight
      }else{
        padding = lineHeight*0.45
      }

      // draw licensee aveerage
      curr_x = curr_x + (totalWidth*thead_heading[1].width)
      if(i == 77){
        doc.text(licensee_ave, (curr_x + 5), (curr_y + padding), {'width': (totalWidth*thead_heading[2].width) - 10, 'align': 'center'});
      }else{
        doc.text(licensee_ave + ' said yes', (curr_x + 5), (curr_y + padding), {'width': (totalWidth*thead_heading[2].width) - 10, 'align': 'center'});
      }

      // draw industry average
      curr_x = curr_x + (totalWidth*thead_heading[2].width)

      doc.text(industry_ave, (curr_x + 5), (curr_y + padding), {'width': (totalWidth*thead_heading[3].width) - 10, 'align': 'center'});

      // reset x-coord to the leftmost writable page of the page
      curr_x = opts.page.margin.left;

      curr_y = curr_y + box_height

      // @TODO, this should be dynamically allocated
      switch(i){
        case 73:
          question = 'If yes, explain how your problem was resolved:'
          box_height = lineHeight*6
        break;
        case 75:
          question = 'Why do you say that?'
          box_height = lineHeight*6
        break;
        case 77:
          question = 'What does the planner have to do to gain your trust?'
          box_height = lineHeight*6
        break;
      }

      // text explanation
      c_addr = XLSX.utils.encode_cell({'c': i+1, 'r': Row})
      if(worksheet[c_addr]){
        answer = worksheet[c_addr].v
      }else{
        answer = ''
      }

      // dynamically allocation of table cell height
      answer = '\“' + answer + '\”'
      doc_width = doc.widthOfString(answer)
      lines = lines_func(doc_width, (totalWidth - 20))
      box_height = doc.currentLineHeight() * (lines + 2.5)
      doc.rect(curr_x, curr_y, totalWidth, box_height)
      doc.fillAndStroke('#dedede', '#000')
      doc.fillColor('#000')

      doc.font(opts.font.weight.boldItalic)
      doc.text(question, (curr_x + 5), (curr_y + (lineHeight*0.45)), {'width': totalWidth - 5, 'align': 'left'})

      doc.font(opts.font.weight.italic)
      doc.text(answer, (curr_x + 5), curr_y + (lineHeight*1.5), {'width': totalWidth - 5, 'align': 'left'})

      curr_x = opts.page.margin.left;
      curr_y = curr_y + box_height
      k++;
    }

   //doc.flushPages()


    // Environment page
    doc.addPage(pageOptions)

    curr_x = opts.page.margin.left;
    curr_y = opts.page.margin.top;

    box_par_width = opts.page.width - (opts.page.margin.left + opts.page.margin.right)
    parOpts = {'width': box_par_width, 'align': 'left'}

    doc.fontSize(20)
    doc.font(opts.font.weight.bold)
    doc.text('ENVIRONMENT', curr_x, curr_y, parOpts)

    doc.fontSize(opts.font.size.normal)
    doc.font(opts.font.weight.normal)
    lineHeight = doc.currentLineHeight()

    curr_y = doc.y + lineHeight

    data = 'Environment measures the intangible and tangible aspects of client-planner experience.  In addition to the physical environment, this category looks at dimensions such as the style and manner of the planner and how easy it was to make an appointment. Historically, this has been a high scoring measure for most participants and is largely considered to be a hygiene factor.'

    doc.font(opts.font.weight.bold)
    doc.text(data.slice(0, 11), curr_x, curr_y, {'width': box_par_width, 'align': 'left', 'continued': true})
    doc.font(opts.font.weight.normal)
    doc.text(data.slice(11))

    doc.lineWidth(0.3)

    curr_x = opts.page.margin.left
    curr_y = doc.y + lineHeight
    box_height = lineHeight*2

    table_heading = ['Your score', licensee_name, 'Industry Average']
    // this should load the data
    table_row = []
    table_row[0] = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 93, 'r': Row}) ].v )
    table_row[1] = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 154, 'r': Row}) ].v )
    table_row[2] = '91'

    doc.font(opts.font.weight.bold)
    for(var i = 0; i<table_heading.length; i++){
      doc.rect(curr_x, curr_y, (totalWidth/3), box_height)
      doc.fillAndStroke(header_fill_style, '#000')
      doc.fillColor(table_header_color)
      doc.text(table_heading[i], (curr_x + 2.5), (curr_y+(box_height*0.30)), {'width': (totalWidth/3) - 5, 'align': 'center'})
      curr_x = curr_x + (totalWidth/3)
    }

    // position relative to the table header
    curr_y = curr_y + box_height

    doc.fontSize(36)
    lineHeight = doc.currentLineHeight()
    box_height = lineHeight*2

    curr_x = opts.page.margin.left
    for(var i = 0; i<table_row.length; i++){
      doc.rect(curr_x, curr_y, (totalWidth/3), box_height)
      doc.fillAndStroke('#fff', '#000')
      doc.fillColor('#000')
      doc.text(table_row[i], curr_x, (curr_y+(lineHeight*0.50)), {'width': (totalWidth/3), 'align': 'center'})

      curr_x = curr_x + (totalWidth/3)
    }

    doc.fontSize(opts.font.size.normal)
    doc.font(opts.font.weight.bold)
    lineHeight = doc.currentLineHeight()

    curr_x = opts.page.margin.left
    curr_y = doc.y + (lineHeight*3)

    doc.font(opts.font.weight.bold)
    doc.text('The shopper was also asked to rate the planner on the following:', curr_x, curr_y, parOpts)

    curr_y = doc.y + lineHeight

    // 80 = EasyTalk, 81 = SocCom, 82 = Friendly, 83 = OnTime, 84 = ProfDressV, 85 = StyleApp
    for(var i=80; i<=85; i++){
      q_key = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 0}) ].v
//      question = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 1}) ].v.replace(' (Answer)', '').replace(' (Text)', '').replace(' (Value)', '')
      switch(i){
        case 80:
          question = 'Easy to talk to'
          break;
        case 81:
          question = 'Socially comfortable'
          break;
        case 82:
          question = 'Friendly'
          break;
        case 83:
          question = 'On time'
          break;
        case 84:
          question = 'Professionalism of dress standards'
          break;
        case 85:
          question = 'Style and manner of approach'
          break;
      }

      doc.font(opts.font.weight.boldItalic)
      doc.text(question, curr_x, curr_y, {'align': 'left', 'width': box_par_width})
      doc.font(opts.font.weight.normal)

      curr_y = doc.y
      doc.image('resources/' + output_dir + '/' + q_key + '.png', curr_x, curr_y, {'scale': 0.70})

      // images are 50px tall
      curr_y = curr_y + (lineHeight*0.45) + 40
    }

   //doc.flushPages()

    // Environment page 2
    doc.addPage(pageOptions)

    curr_x = opts.page.margin.left;
    curr_y = opts.page.margin.top;

    box_par_width = opts.page.width - (opts.page.margin.left + opts.page.margin.right)
    parOpts = {'width': box_par_width, 'align': 'left'}

    doc.fontSize(opts.font.size.normal)
    doc.font(opts.font.weight.normal)
    lineHeight = doc.currentLineHeight()

    // 86 = LongAns, 87 = PeopSpeak, 88 = ContactFP 
    for(var i=86; i<=88; i++){
      q_key = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 0}) ].v
//      question = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 1}) ].v.replace(' (Answer)', '').replace(' (Text)', '')
      answer = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': Row}) ].v
      switch(i){
        case 86:
          question = 'How long did it take to answer the first call you made to this organisation?'
          break;
        case 87:
          question = 'How many people did you have to speak to before you could make an appointment?'
          break;
        case 88:
          question = 'How many times did you have to make contact with the planner or organisation to get an appointment?'
          break;
        default:
          question = ''
          break;
      }

      doc.font(opts.font.weight.boldItalic)
      doc.text(question, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})

      doc.font(opts.font.weight.normal)
      //get max width for the questions
      max_width = doc.widthOfString('Your Shopper: ') + 30
      curr_y = doc.y + lineHeight

      doc.text('Your Shopper:', curr_x, curr_y, {'width': max_width, 'align': 'left'})
      doc.font(opts.font.weight.bold)
      doc.text(answer, (doc.x + max_width), curr_y, {'width': (box_par_width - max_width), 'align': 'left'})

      curr_y = doc.y + lineHeight
      doc.font(opts.font.weight.normal)
      doc.image(licensee_images + '/' + q_key + '.png', curr_x, curr_y, {'scale': 0.1275})

      curr_y = doc.y + (lineHeight*13);
    }

   //doc.flushPages()

    // Environment page 3
    doc.addPage(pageOptions)

    curr_x = opts.page.margin.left;
    curr_y = opts.page.margin.top;

    box_par_width = opts.page.width - (opts.page.margin.left + opts.page.margin.right)
    parOpts = {'width': box_par_width, 'align': 'left'}

    doc.fontSize(opts.font.size.normal)
    doc.font(opts.font.weight.normal)
    lineHeight = doc.currentLineHeight()

    // 89 = Helpful, 90 = EasyApp
    for(var i=89; i<=90; i++){
      q_key = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 0}) ].v
//      question = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 1}) ].v.replace(' (Answer)', '').replace(' (Text)', '').replace(' (Value)', '')
      switch(i){
        case 89:
          question = 'Rate the helpfulness of the person you spoke to in the planner’s office when you rang.'
          break;
        case 90:
          question = 'Overall, how easy was it to make an appointment?'
          break;
        default:
          question = ''
          break;
      }

      doc.font(opts.font.weight.boldItalic)
      doc.text(question, curr_x, curr_y, {'align': 'left', 'width': box_par_width})
      doc.font(opts.font.weight.normal)

      curr_y = doc.y
      doc.image('resources/' + output_dir + '/' + q_key + '.png', curr_x, curr_y, {'scale': 0.70})

      // images are 50px tall
      curr_y = curr_y + (lineHeight*0.45) + 50
    }

    doc.fontSize(opts.font.size.normal)
    doc.font(opts.font.weight.bold)
    lineHeight = doc.currentLineHeight()

    curr_y = doc.y + (lineHeight*3)
    doc.font(opts.font.weight.bold)
    doc.text('The shopper was also asked:', curr_x, curr_y, parOpts)
    doc.font(opts.font.weight.normal)

    curr_y = doc.y + lineHeight

    doc.lineWidth(0.3)

    // track the y-coords for table heading
    // the table headers
    thead_heading = [];
    thead_heading[0] = {'width': 0.46, 'data': ''};
    thead_heading[1] = {'width': 0.18, 'data': 'Your Shopper'};
    thead_heading[2] = {'width': 0.18, 'data': licensee_name};
    thead_heading[3] = {'width': 0.18, 'data': 'Industry Average'};

    // lineJoin() cause a subtle bug
    doc.lineCap('square');

    // draw the table headings
    doc.font(opts.font.weight.bold)
    doc.fontSize(opts.font.size.normal)

    writable_width = (totalWidth*thead_heading[2].width) - 5
    doc_width = doc.widthOfString(thead_heading[2].data)
    answer_lines = lines_func(doc_width, writable_width)

    writable_width = (totalWidth*thead_heading[3].width) - 5
    doc_width = doc.widthOfString('Industry Average')
    licensee_lines = lines_func(doc_width, writable_width)

    max_line = Math.max(answer_lines, licensee_lines) + 1
    box_height = lineHeight*max_line

    curr_y = doc.y + lineHeight;
    for(var i=0; i<thead_heading.length; i++){

      if(i == 2){
        padding = ((max_line - answer_lines) * lineHeight) * 0.50
      }else{
        padding = lineHeight * 0.40;
      }

      doc.rect(curr_x, curr_y, (totalWidth*thead_heading[i].width), box_height)
      doc.fillAndStroke(header_fill_style, '#000')

      doc.fillColor(table_header_color)
      doc.text(thead_heading[i].data, (curr_x + 2.5), (curr_y + padding), {'align': 'center', 'width': (totalWidth*thead_heading[i].width) - 5})

      curr_x = curr_x + (totalWidth*thead_heading[i].width)
    }

    doc.font(opts.font.weight.normal)

    // resets back to orig x-coord
    curr_x = opts.page.margin.left;

    // draw table rows
    // iterate the table row headings
    // add padding between rect box boundary and the text
    // each row is positioned using tableOpts['trow_height']*(i+1)

    curr_y = curr_y + box_height

    // @todo put it in the top
    scoreVal = []

//    question = worksheet[ XLSX.utils.encode_cell({'c': 94, 'r': 1}) ].v.replace(' (Text)', '')
    answer = worksheet[ XLSX.utils.encode_cell({'c': 91, 'r': Row}) ].v
    licensee_ave = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 163, 'r': Row}) ].v ) + '%'
    scoreVal[0] = {'key': 'ExtBldg', 'question': 'How did you perceive the office building’s exterior appearance?', 'answer': answer, 'licensee': licensee_ave, 'industry': '94%'}

    answer = worksheet[ XLSX.utils.encode_cell({'c': 92, 'r': Row}) ].v
    licensee_ave = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 164, 'r': Row}) ].v ) + '%'
    scoreVal[1] = {'key': 'EnviBldg', 'question': 'How did you perceive the interior environment of the planner’s office?', 'answer': answer, 'licensee': licensee_ave, 'industry': '94%'}

    // reset box_height for table rows
    box_height = (lineHeight*4)

    for(var i = 0; i < scoreVal.length; i++){
      var style = (i % 2) == 1 ? '#fafafa' : '#dedede';

      // draw the bounding rect first
      for(var j=0; j<thead_heading.length; j++){
        doc.rect(curr_x, curr_y, (totalWidth*thead_heading[j].width), box_height)
        doc.fillAndStroke(style, '#000')

        curr_x += (totalWidth*thead_heading[j].width)
      }

      // resets back to orig x-coord
      curr_x = opts.page.margin.left

      // draw the text within the bounding rect
      doc.fillColor('#000')
      switch(scoreVal[i].key){
        case 'EnviBldg':
          padding = lineHeight*0.30
          break;
        default:
          padding = lineHeight *0.90
          break;
      }

      doc.font(opts.font.weight.boldItalic)
      doc.text(scoreVal[i].question, (curr_x + 2.5), (curr_y + padding), {'width': (totalWidth*thead_heading[0].width) - 5, 'align': 'left'});
      doc.font(opts.font.weight.normal)

      switch(scoreVal[i].answer.toLowerCase()){
        case 'no answer':
          padding = lineHeight*1.5
          break;

        case 'exceeded my expectations':
          padding = lineHeight*0.30
          break;
        default:
          padding = lineHeight*0.90
          break;
      }

      // draw the self rating
      curr_x = curr_x + (totalWidth*thead_heading[0].width)
      doc.text(scoreVal[i].answer, (curr_x + 2.5), (curr_y + padding), {'width': (totalWidth*thead_heading[1].width) - 5, 'align': 'center'});
      doc.font(opts.font.weight.normal)

      padding = lineHeight*0.30

      // draw licensee average
      padding = lineHeight*0.275
      curr_x = curr_x+(totalWidth*thead_heading[1].width)
      doc.text(scoreVal[i].licensee + ' met or exceeded expectations', (curr_x + 2.5), (curr_y + padding), {'width': (totalWidth*thead_heading[2].width) - 5, 'align': 'center'});

      // draw industry average
      padding = lineHeight*0.275
      curr_x = curr_x+(totalWidth*thead_heading[2].width)
      doc.text(scoreVal[i].industry + ' met or exceeded expectations', (curr_x + 2.5), (curr_y + padding), {'width': (totalWidth*thead_heading[3].width) - 5, 'align': 'center'});

      // reset x-coord to the leftmost writable page of the page
      curr_x = opts.page.margin.left;
      curr_y = curr_y + box_height
    }

    // Follow up page
    doc.addPage(pageOptions)

    curr_x = opts.page.margin.left;
    curr_y = opts.page.margin.top;

    box_par_width = opts.page.width - (opts.page.margin.left + opts.page.margin.right)
    parOpts = {'width': box_par_width, 'align': 'left'}

    doc.fontSize(20)
    doc.font(opts.font.weight.bold)
    doc.text('FOLLOW UP', curr_x, curr_y, parOpts)

    doc.fontSize(opts.font.size.normal)
    doc.font(opts.font.weight.normal)
    lineHeight = doc.currentLineHeight()

    curr_y = doc.y + lineHeight
    data = 'One of the recurring findings of this study is poor follow up rates by planners after the first meeting. Best practice suggests that all customers should be followed up within two days of the meeting, ideally by phone.  Research shows clearly that the likelihood of a shopper using or recommending a planner is significantly higher when they are followed up.'
    doc.text(data, curr_x, curr_y, parOpts)

    doc.lineWidth(0.3)

    // track the y-coords for table heading
    // the table headers
    var thead_heading = [];
    thead_heading[0] = {'width': 0.55, 'data': ''};
    thead_heading[1] = {'width': 0.15, 'data': 'Your Shopper'};
    thead_heading[2] = {'width': 0.15, 'data': licensee_name};
    thead_heading[3] = {'width': 0.15, 'data': 'Industry Average'};

    // lineJoin() cause a subtle bug
    doc.lineCap('square');

    // draw the table headings
    doc.font(opts.font.weight.bold)
    doc.fontSize(opts.font.size.normal)

    lineHeight = doc.currentLineHeight()

    curr_x = opts.page.margin.left;
    curr_y = doc.y + lineHeight;

    writable_width = (totalWidth*thead_heading[2].width) - 5
    doc_width = doc.widthOfString(thead_heading[2].data)
    answer_lines = lines_func(doc_width, writable_width)

    writable_width = (totalWidth*thead_heading[3].width) - 5
    doc_width = doc.widthOfString('Industry Average')
    licensee_lines = lines_func(doc_width, writable_width)

    max_line = Math.max(answer_lines, licensee_lines) + 1
    box_height = lineHeight*max_line

    for(var i=0; i<thead_heading.length; i++){
      doc.rect(curr_x, curr_y, (totalWidth*thead_heading[i].width), box_height)
      doc.fillAndStroke(header_fill_style, '#000')

      doc.fillColor(table_header_color)

      if( i==2 ){
        padding = ((max_line - answer_lines) * lineHeight) * 0.50
      }else{
        padding = lineHeight * 0.40
      }

      doc.text(thead_heading[i].data, (curr_x + 2.5), (curr_y + padding), {'align': 'center', 'width': (totalWidth*thead_heading[i].width) - 5})

      curr_x = curr_x + (totalWidth*thead_heading[i].width)
    }

    doc.font(opts.font.weight.normal)

    // resets back to orig x-coord
    curr_x = opts.page.margin.left;

    // draw table rows
    // iterate the table row headings
    // add padding between rect box boundary and the text
    // each row is positioned using tableOpts['trow_height']*(i+1)

    curr_y = curr_y + box_height

    // @todo put it in the top
    scoreVal = []

//    question = worksheet[ XLSX.utils.encode_cell({'c': 94, 'r': 1}) ].v.replace(' (Text)', '')
    answer = worksheet[ XLSX.utils.encode_cell({'c': 94, 'r': Row}) ].v
    licensee_ave = norm_func( worksheet[ XLSX.utils.encode_cell({'c': 165, 'r': Row}) ].v ) + '%'
    scoreVal[0] = {'key': 'FollowUp', 'question': 'After the first meeting were you followed up?', 'answer': answer, 'licensee': licensee_ave, 'industry': '42%'}

    for(var i = 0; i < scoreVal.length; i++){
      var style = (i % 2) == 1 ? '#fafafa' : '#dedede';

      // draw the bounding rect first
      for(var j=0; j<thead_heading.length; j++){
        doc.rect(curr_x, curr_y, (totalWidth*thead_heading[j].width), box_height)
        doc.fillAndStroke(style, '#000')

        curr_x += (totalWidth*thead_heading[j].width)
      }

      // resets back to orig x-coord
      curr_x = opts.page.margin.left

      writable_width = (totalWidth*thead_heading[0].width) - 5
      doc_width = doc.widthOfString(scoreVal[i].question)
      question_lines = lines_func(doc_width, writable_width)

      writable_width = (totalWidth*thead_heading[2].width) - 5
      doc_width = doc.widthOfString(scoreVal[i].answer)
      answer_lines = lines_func(doc_width, writable_width)

      writable_width = (totalWidth*thead_heading[3].width) - 5
      doc_width = doc.widthOfString(scoreVal[i].licensee + ' said yes')
      licensee_lines = lines_func(doc_width, writable_width)

      writable_width = (totalWidth*thead_heading[3].width) - 5
      doc_width = doc.widthOfString(scoreVal[i].industry + ' said yes')
      industry_lines = lines_func(doc_width, writable_width)

      max_line = Math.max(question_lines, answer_lines, licensee_lines, industry_lines) + 1
      box_height = lineHeight*max_line

      padding = ((max_line - question_lines) * lineHeight) * 0.45
      // draw the text within the bounding rect
      doc.fillColor('#000')
      doc.font(opts.font.weight.boldItalic)
      doc.text(scoreVal[i].question, (curr_x + 2.5), (curr_y + padding), {'width': (totalWidth*thead_heading[0].width) - 5, 'align': 'left'});
      doc.font(opts.font.weight.normal)

      // draw the self rating
      padding = ((max_line - answer_lines) * lineHeight) * 0.45
      curr_x = curr_x + (totalWidth*thead_heading[0].width)
      doc.text(scoreVal[i].answer, (curr_x + 2.5), (curr_y + padding), {'width': (totalWidth*thead_heading[1].width) - 5, 'align': 'center'});

      // draw licensee average
      padding = ((max_line - licensee_lines) * lineHeight) * 0.45
      curr_x = curr_x + (totalWidth*thead_heading[1].width)
      doc.text(scoreVal[i].licensee + ' said yes', (curr_x + 2.5), (curr_y + padding), {'width': (totalWidth*thead_heading[2].width) - 5, 'align': 'center'});

      // draw industry average
      padding = ((max_line - industry_lines) * lineHeight) * 0.45
      curr_x = curr_x + (totalWidth*thead_heading[2].width)
      doc.text(scoreVal[i].industry + ' said yes', (curr_x + 2.5), (curr_y + padding), {'width': (totalWidth*thead_heading[3].width) - 5, 'align': 'center'});

      // reset x-coord to the leftmost writable page of the page
      curr_x = opts.page.margin.left;
      curr_y = curr_y + box_height
    }

    curr_y = doc.y + (lineHeight*4)

    // 95 = DaysFollow, 96 = HowFollow
    for(var i=95; i<=96; i++){
      q_key = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 0}) ].v
      question = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': 1}) ].v.replace(' (Answer)', '').replace(' (Text)', '')
      answer = worksheet[ XLSX.utils.encode_cell({'c': i, 'r': Row}) ].v

      doc.font(opts.font.weight.boldItalic)
      doc.text(question, curr_x, curr_y, {'width': box_par_width, 'align': 'left'})
      doc.font(opts.font.weight.normal)
      //get max width for the questions
      max_width = doc.widthOfString('Your Shopper: ') + 30
      curr_y = doc.y + lineHeight

      doc.text('Your Shopper:', curr_x, curr_y, {'width': max_width, 'align': 'left'})
      doc.font(opts.font.weight.bold)
      doc.text(answer, (doc.x + max_width), curr_y, {'width': (box_par_width - max_width), 'align': 'left'})

      curr_y = doc.y + lineHeight
      doc.font(opts.font.weight.normal)
      doc.image(licensee_images + '/' + q_key + '.png', curr_x, curr_y, {'scale': 0.1275})

      curr_y = doc.y + (lineHeight*13);
    }

    // Pipe it's output somewhere, like to a file or HTTP response
    // See below for browser usage
    //doc.pipe(res)
    doc.pipe(fs.createWriteStream(target_output_dir + '/' + licensee_name + ' - ' + plannerName + '.pdf'))

    // Finalize PDF file
    doc.end()

  }

  // force the requesting resource to download the pdf
  //res.setHeader('Content-disposition', 'attachment; filename=cer-report.pdf');


  //res.send('processed ranges:' + JSON.stringify(range));
//});
//
//app.listen(8000, function () {
//  console.log('Application listening on port 8000!');
//});
