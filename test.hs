module Main where

import Text.XML.SpreadsheetML.Types
import Text.XML.SpreadsheetML.Writer ( showSpreadsheet )
import Text.XML.SpreadsheetML.Builder

import System.IO

cells1 = [ [string "Quantity", string "Multiplier", string "Product"]
         , [number 1,          number 0.9,          formula "=RC[-2]*RC[-1]"]
         , [number 10,         number 1.1,          formula "=RC[-2]*RC[-1]"]
         , [number 12,         number 0.2,          formula "=RC[-2]*RC[-1]"]
         ]
worksheet1 = mkWorksheet (Name "Quantity Product Sheet") (tableFromCells cells1)
cells2 = [ [string "Quantity1", string "Quantity2", string "Sum"]
         , [number 1,          number 100,          formula "=RC[-2]+RC[-1]"]
         , [number 10,         number 201,          formula "=RC[-2]+RC[-1]"]
         , [number 12,         number 45,           formula "=RC[-2]+RC[-1]"]
         ]
worksheet2 = mkWorksheet (Name "Quantity Sum Sheet") (tableFromCells cells2)
workbook = mkWorkbook [worksheet1, worksheet2]

main = writeFile "test.xls" (showSpreadsheet workbook)
