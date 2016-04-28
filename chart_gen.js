/* vim: sw=2 ts=2 expandtab */
const exec = require('child_process').exec;

var express = require('express');
var app = express();

var PDFDocument = require('pdfkit');

var fs = require('fs');
var XLSX = require('xlsx');

var command_args = process.argv

// start reading the XLSX file
var workbook = XLSX.readFile('WW\ 2016\ Overall\ Individual\ Adviser\ Report\ Data\ \(Additional\ Reports\)\ v1.3.xlsx');

//excel file may contain multiple sheets, 1 sheet = 1 client
var sheet_name_list = workbook.SheetNames;

// get the range for rows and columns
//var ranges = workbook.Sheets[sheet_name_list[0]]['!ref'];
var ranges = workbook.Sheets['Individual Report']['!ref'];
var range = XLSX.utils.decode_range(ranges);

var target_licensee = 'St. George Financial Planning'
target_licensee = target_licensee.toLowerCase()

console.log(range)
//var worksheet = workbook.Sheets[sheet_name_list[0]]
var worksheet = workbook.Sheets['Individual Report']

var args, review_number, output_dir, target_file

// read the row data one by one
// start of usable row is 2, 0-index rule
//  for(var Row = 2; Row <= range.e.r; Row++) {

// this structure contains the column location where we should be fetching the values for
// planner score, licensee score, industry score
var input_cols = []
input_cols.push({'q_key': 'AbilityDemo', 'p_score': 28, 'l_score': 113, 'i_score': 166})
input_cols.push({'q_key': 'AbilityExp', 'p_score': 29, 'l_score': 114, 'i_score': 167})
input_cols.push({'q_key': 'ClearEasy', 'p_score': 31, 'l_score': 115, 'i_score': 168})
input_cols.push({'q_key': 'AbDemoEff', 'p_score': 32, 'l_score': 116, 'i_score': 169})
input_cols.push({'q_key': 'ConvReco', 'p_score': 47, 'l_score': 130, 'i_score': 170})
input_cols.push({'q_key': 'FeesPay', 'p_score': 48, 'l_score': 131, 'i_score': 171})
input_cols.push({'q_key': 'PlanSrvcs', 'p_score': 49, 'l_score': 132, 'i_score': 172})
input_cols.push({'q_key': 'ListenSkill', 'p_score': 51, 'l_score': 134, 'i_score': 173})
input_cols.push({'q_key': 'Goals', 'p_score': 52, 'l_score': 135, 'i_score': 174})
input_cols.push({'q_key': 'DemoGoals', 'p_score': 53, 'l_score': 136, 'i_score': 175})
input_cols.push({'q_key': 'ReadFact', 'p_score': 194, 'l_score': 137, 'i_score': 176})
input_cols.push({'q_key': 'WellPrep', 'p_score': 55, 'l_score': 138, 'i_score': 177})
input_cols.push({'q_key': 'M_2ndMeet', 'p_score': 57, 'l_score': 140, 'i_score': 178})
input_cols.push({'q_key': 'RecoP', 'p_score': 59, 'l_score': 141, 'i_score': 179})
input_cols.push({'q_key': 'Keen', 'p_score': 66, 'l_score': 145, 'i_score': 180})
input_cols.push({'q_key': 'Gimpress', 'p_score': 68, 'l_score': 146, 'i_score': 181})
input_cols.push({'q_key': 'Influence', 'p_score': 69, 'l_score': 147, 'i_score': 182})
input_cols.push({'q_key': 'Enthuse', 'p_score': 70, 'l_score': 148, 'i_score': 183})
input_cols.push({'q_key': 'Reltn', 'p_score': 71, 'l_score': 149, 'i_score': 184})
input_cols.push({'q_key': 'Rapprt', 'p_score': 72, 'l_score': 150, 'i_score': 185})
input_cols.push({'q_key': 'EasyTalk', 'p_score': 80, 'l_score': 155, 'i_score': 186})
input_cols.push({'q_key': 'SocCom', 'p_score': 81, 'l_score': 156, 'i_score': 187})
input_cols.push({'q_key': 'Friendly', 'p_score': 82, 'l_score': 157, 'i_score': 188})
input_cols.push({'q_key': 'OnTime', 'p_score': 83, 'l_score': 158, 'i_score': 189})
input_cols.push({'q_key': 'ProfDressV', 'p_score': 84, 'l_score': 159, 'i_score': 190})
input_cols.push({'q_key': 'StyleApp', 'p_score': 85, 'l_score': 160, 'i_score': 191})
input_cols.push({'q_key': 'Helpful', 'p_score': 89, 'l_score': 161, 'i_score': 192})
input_cols.push({'q_key': 'EasyApp', 'p_score': 90, 'l_score': 162, 'i_score': 193})

//start_row = range.s.r
//end_row = range.e.r
col_row = command_args[2]
//start_row = col_row - 1
//end_row = start_row

start_row = 2
end_row = range.e.r

var number = ''
var licensee_full_text

for(var Row = start_row; Row <= end_row; Row++) {
//  q_key = worksheet[ XLSX.utils.encode_cell({'c': 28, 'r': 0}) ].v

  if(worksheet[ XLSX.utils.encode_cell({'c': 0, 'r': Row}) ]){
    number = worksheet[ XLSX.utils.encode_cell({'c': 0, 'r': Row}) ].v
    licensee_full_text = worksheet[ XLSX.utils.encode_cell({'c': 5, 'r': Row}) ].v
    // limit to 4 digit filename
  //  number = ("0000" + number).substr(-4,4)

    if(licensee_full_text.toLowerCase() == 'commonwealth financial planning' || licensee_full_text.toLowerCase() == 'financial wisdom' || licensee_full_text.toLowerCase() == 'count financial'){

      output_dir = './resources/' + number

      console.log('PROCESSING: ' + number)
      for(i_col in input_cols){
        q_key = input_cols[i_col].q_key

        console.log('\tkey: ' + q_key)

        if(worksheet[ XLSX.utils.encode_cell({'c': input_cols[i_col].p_score, 'r': Row}) ]){
          console.log('\t' + JSON.stringify( worksheet[ XLSX.utils.encode_cell({'c': input_cols[i_col].p_score, 'r': Row}) ] ))
        }else{
          console.log('\tERROR: ' + q_key)
        }

        if(worksheet[ XLSX.utils.encode_cell({'c': input_cols[i_col].p_score, 'r': Row}) ]){
          planner_score = worksheet[ XLSX.utils.encode_cell({'c': input_cols[i_col].p_score, 'r': Row}) ].v
        }

        if(planner_score == 'No applicable'){
          planner_score = -1
        }

        if(planner_score == 'No answer'){
          planner_score = -1
        }

        if(planner_score == 'Not applicable'){
          planner_score = -1
        }

        if(worksheet[ XLSX.utils.encode_cell({'c': input_cols[i_col].l_score, 'r': Row}) ]){
          licensee_score = worksheet[ XLSX.utils.encode_cell({'c': input_cols[i_col].l_score, 'r': Row}) ].v
        }else{
          licensee_score = 0
        }

        if(worksheet[ XLSX.utils.encode_cell({'c': input_cols[i_col].i_score, 'r': Row}) ]){
          industry_Score = worksheet[ XLSX.utils.encode_cell({'c': input_cols[i_col].i_score, 'r': Row}) ].v
        }else{
          industry_Score = 0
        }

        target_file = output_dir + '/' + q_key + '.png'
        args = planner_score + ' ' + licensee_score + ' ' + industry_Score + ' ' + target_file

        // let the child process manage the creation of images and the stdout
        exec('./test_bullet.php ' + args,
          function (error, stdout, stderr) {
            if (error !== null) {
              console.log('exec error: ' + error);
            }

            fs.appendFile( 'url_list.log', stdout + '\n\n')
            for(var i=1;i<=10000;i++){
              // do nothing
            }
        });

      }
    }
  }
}
