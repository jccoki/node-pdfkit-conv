/* vim: sw=2 ts=2 expandtab */
const exec = require('child_process').exec;

var express = require('express');
var app = express();

var PDFDocument = require('pdfkit');

var fs = require('fs');
var XLSX = require('xlsx');

// start reading the XLSX file
var workbook = XLSX.readFile('WW\ 2016\ Overall\ Individual\ Adviser\ Report\ Data\ v1.8.xlsx');

//excel file may contain multiple sheets, 1 sheet = 1 client
var sheet_name_list = workbook.SheetNames;

// get the range for rows and columns
//var ranges = workbook.Sheets[sheet_name_list[0]]['!ref'];
var ranges = workbook.Sheets['Individual Report']['!ref'];
var range = XLSX.utils.decode_range(ranges);

console.log(range)
//var worksheet = workbook.Sheets[sheet_name_list[0]]
var worksheet = workbook.Sheets['Individual Report']

//{ s: { c: 0, r: 0 }, e: { c: 192, r: 105 } }
//for(var Row = 2; Row <= range.e.r; Row++) {
//for(var Row = 2; Row <= 2; Row++) {
  for(var Col = range.s.c; Col <= range.e.c; Col++){
    q_key = worksheet[ XLSX.utils.encode_cell({'c': Col, 'r': 0}) ].v
    fs.appendFile( 'column.lst',  q_key + ' -> ' + Col + '\n')
  }
//}
