module Main where

import Text.XML.SpreadsheetML.Types
import Text.XML.SpreadsheetML.Writer

wsName = Name "Worksheet1"
dp = DocumentProperties Nothing Nothing Nothing Nothing Nothing (Just "GHCi") Nothing
cells = [ Cell (Just (Number 1.0)) Nothing Nothing Nothing Nothing, Cell (Just (StringType "Blah!")) Nothing Nothing Nothing Nothing ]
row = Row cells Nothing Nothing Nothing Nothing Nothing Nothing
table = Table [] [row] Nothing Nothing Nothing Nothing Nothing Nothing Nothing Nothing
worksheet = Worksheet (Just table) wsName
workbook = Workbook (Just dp) [worksheet]

main = putStrLn $ showSpreadsheet workbook
